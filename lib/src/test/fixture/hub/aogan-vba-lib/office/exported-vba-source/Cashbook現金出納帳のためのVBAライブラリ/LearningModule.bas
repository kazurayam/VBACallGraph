Attribute VB_Name = "LearningModule"
Option Explicit

' VBA���R�[�f�B���O���邤���ŋ^��Ɏv�������Ƃ��������邽��
' ���낢�뎎�����B���̃R�[�h�������ɔz�u����B

Sub ListWorksheetsInAoganCashbook()
    ' �X����Ȉ��̌����o�[���̃��[�N�u�b�N���J���A
    ' ���̂Ȃ��ɂ��郏�[�N�V�[�g����Debug.Print����
    Dim wb As Workbook
    Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Dim i As Long
    Call KzCls
    For i = 1 To wb.Worksheets.Count
        Debug.Print wb.Worksheets(i).Name
    Next i
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
End Sub



