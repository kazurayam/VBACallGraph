Attribute VB_Name = "KzMetaprogramming"
Option Explicit


Public Sub KzProcedureList(ByVal wb As Workbook)

    Dim dicProcInfo As New Dictionary
    Dim i As Long
  
    '�u�b�N�̑S���W���[��������
    With wb.VBProject
        For i = 1 To .VBComponents.Count
            Call getCodeModule.getCodeModule(dicProcInfo, wb, .VBComponents(i).Name)
        Next
    End With
  
    '�o�͐�Ƃ��Ẵ��[�N�V�[�g����������
    Dim sheetName As String: sheetName = "�v���V�[�W���ꗗ"
    Dim r As Boolean
    r = KzCreateWorksheetInWorkbook(wb, sheetName)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    
    '�v���V�[�W���[�̏����V�[�g�ɏo�͂���
    Dim v
    With ws
        .Cells.Clear
        .Range("A1:G1").Value = Array(wb.Name, "���W���[��", "�X�R�[�v", "���", "�s�ʒu", "�\�[�X", "�R�����g")
        .Range("A1:G1").Interior.Color = RGB(200, 200, 200) ' �w�i�F���O���[��
        i = 2
        For Each v In dicProcInfo.Items
            .Cells(i, 1) = v.ProcName
            .Cells(i, 2) = v.ModName
            .Cells(i, 3) = v.Scope
            ' .Cells(i, 4) = v.ProcKindName
            .Cells(i, 4) = IIf(InStr(1, LCase(v.Source), " function ") > 0, "Function", "Sub")
            .Cells(i, 5) = v.LineNo
            .Cells(i, 6) = v.Source
            .Cells(i, 7) = "'" & v.Comment
            i = i + 1
        Next
        Cells.EntireRow.AutoFit
        Cells.EntireColumn.AutoFit
    End With

    '�V�[�g�̍s���v���V�[�W�����̏����Ń\�[�g����
    With ws.Sort
        With .SortFields
            .Clear
            .Add key:=ws.Range("A1"), Order:=xlAscending
        End With
        .SetRange ws.Range(Cells(1, 1), Cells(i, 7))
        .Header = xlYes
        .Apply
    End With
    
    '�s�̍������������߂���
    ws.Rows.AutoFit
    
    Set dicProcInfo = Nothing
End Sub

