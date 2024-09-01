Attribute VB_Name = "BbWorksheet"
Option Explicit

'KzWorksheet


' �w�肳�ꂽ���[�N�u�b�N�̂Ȃ��Ɏw�肳�ꂽ���O�̃V�[�g�����݂��Ă�����True��Ԃ�
Public Function IsWorksheetPresentInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            flg = True
            Exit For
        End If
    Next ws
    IsWorksheetPresentInWorkbook = flg
End Function


' �w�肳�ꂽ���[�N�u�b�N�̂Ȃ��Ɏw�肳�ꂽ���̃V�[�g�����݂��Ȃ���Βǉ�����
' �ǉ������Ƃ���True��Ԃ��B
' �V�[�g�����łɂ������Ȃ�΂Ȃɂ�����False��Ԃ�
Public Function CreateWorksheetInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim flg As Boolean: flg = False
    If Not IsWorksheetPresentInWorkbook(wb, sheetName) Then
        Dim ws As Worksheet
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = sheetName
        flg = True
    End If
    CreateWorksheetInWorkbook = flg
End Function


' �w�肳�ꂽ���[�N�u�b�N�̂Ȃ��Ɏw�肳�ꂽ���̃V�[�g�����݂���΍폜����
' �폜�����Ƃ���True��Ԃ��B
' �V�[�g��������΂Ȃɂ�����False��Ԃ�
Public Function DeleteWorksheetInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    '�w�肳�ꂽ�u�b�N�Ɏw�肵���V�[�g�����݂��邩�`�F�b�N
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            '����΃V�[�g���폜����
            Application.DisplayAlerts = False    ' ���b�Z�[�W���\��
            ws.Delete
            Application.DisplayAlerts = True
            flg = True
            Exit For
        End If
    Next ws
    DeleteWorksheetInWorkbook = flg
End Function


' �R�s�[���Ƃ��Ďw�肳�ꂽ���[�N�u�b�N�̃��[�N�V�[�g��
' �R�s�[��Ƃ��Ďw�肳�ꂽ���[�N�u�b�N�̃��[�N�V�[�g�ɃR�s�[����B
' @param sourceWorkbook �R�s�[����Workbook
' @param sourceSheetName �R�s�[����Worksheet�̖��O
' @param targetWorkbook �R�s�[���Workbook
' @param targetSheetName �R�s�[���Worksheet�̖��O
'
' sourceWorkbook��sourceSheetName�Ŏ�����郏�[�N�V�[�g�����݂��Ă��邱�Ƃ��K�v�B�����Ȃ���΃G���[�ɂȂ�B
'
' targetWorkbook�̂Ȃ���targetSheetName�Ŏ�����郏�[�N�V�[�g�����������ꍇ�Ƃ��łɍ݂�ꍇ�Ƃ����肤��B
' �܂��������source�̃V�[�g���R�s�[���邱�Ƃ�targetSheetName�̃��[�N�V�[�g���V�����ł���B
' ���łɍ݂�����targetWorkbook�̂Ȃ��̌Â��V�[�g���폜���āAsource�̃V�[�g���R�s�[����B
' ������sourceWorkbook��targetWorkbook�������ŁA���AsourceSheetName��targetSheetName�������ꍇ��
' �w��~�X������G���[�Ƃ���B
'
Public Sub FetchWorksheetFromWorkbook(ByVal sourceWorkbook As Workbook, _
                                        ByVal sourceSheetName As String, _
                                        ByVal targetWorkbook As Workbook, _
                                        ByVal targetSheetName As String)
'�G���[���N�����Ƃ���ErrorHandler�ɒ���
On Error GoTo ErrorHandler

    'source��target�������ꍇ�̓G���[�Ƃ���
    If sourceWorkbook.path = targetWorkbook.path And sourceSheetName = targetSheetName Then
        Err.Raise Number:=2022, Description:="�������[�N�u�b�N�̓������[�N�V�[�g��source��target�Ɏw�肵�Ă͂����܂���"
    End If
    
    '�W�I�̃��[�N�V�[�g�����łɃ^�[�Q�b�g�̃��[�N�u�b�N�ɂ�������폜����
    If IsWorksheetPresentInWorkbook(targetWorkbook, targetSheetName) Then
        Application.DisplayAlerts = False
        targetWorkbook.Worksheets(targetSheetName).Delete
        Application.DisplayAlerts = True
    End If
    
    '�R�s�[�����[�N�V�[�g�̂��ׂẴZ�����R�s�[����
    '�V�������[�N�V�[�g�Ƃ��ă^�[�Q�b�g�̃��[�N�u�b�N�ɑ}������
    sourceWorkbook.Worksheets(sourceSheetName).Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
    
    '�V�������[�N�V�[�g�̖��O���w�肳�ꂽ�悤�ɕύX����
    ActiveSheet.Name = targetSheetName
        
    
 '��O����
ErrorHandler:
    ' �������G���[���N���Ă����Ȃ�call���ɓ`�d������
    If Err.Number <> 0 Then
        Call BbLog.Info("BbWorksheet", "FetchWorksheetFromWorkbook", "sourceWorkbook : " & sourceWorkbook.FullName)
        Call BbLog.Info("BbWorksheet", "FetchWorksheetFromWorkbook", "sourceSheetName: " & sourceSheetName)
        Call BbLog.Info("BbWorksheet", "FetchWorksheetFromWorkbook", "targetWorkbook : " & targetWorkbook.FullName)
        Call BbLog.Info("BbWorksheet", "FetchWorksheetFromWorkbook", "targetSheetName: " & targetSheetName)
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If

End Sub

