Attribute VB_Name = "KVL�v���V�[�W���[�ꗗ�����"
Option Explicit

Private Sub KVL�v���V�[�W���[�ꗗ�����()

    Dim dicProcInfo As New Dictionary
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim i As Long
  
    '�u�b�N�̑S���W���[��������
    With wb.VBProject
        For i = 1 To .VBComponents.Count
            Call getCodeModule.getCodeModule(dicProcInfo, wb, .VBComponents(i).Name)
        Next
    End With
  
    'Dictionary���V�[�g�ɏo��
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("�v���V�[�W���[�ꗗ")
    Dim v
    With ws
        .Cells.Clear
        .Range("A1:G1").Value = Array("�v���V�[�W���[ in " & ThisWorkbook.Name, "���W���[��", "�X�R�[�v", "���", "�s�ʒu", "�\�[�X", "�R�����g")
        i = 2
        For Each v In dicProcInfo.Items
            .Cells(i, 1) = v.ProcName
            .Cells(i, 2) = v.ModName
            .Cells(i, 3) = v.Scope
            .Cells(i, 4) = v.ProcKindName
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
    
    Set dicProcInfo = Nothing
End Sub
