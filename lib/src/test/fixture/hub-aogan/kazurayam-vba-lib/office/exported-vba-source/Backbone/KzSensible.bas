Attribute VB_Name = "KzSensible"
Option Explicit


Public Sub ExportModules(ByVal wb As Workbook)

    Dim dicProcInfo As New Dictionary
    Dim i As Long
  
    '�u�b�N�̑S���W���[��������
    With wb.VBProject
        For i = 1 To .VBComponents.Count
            Call getCodeModule.getCodeModule(dicProcInfo, wb, .VBComponents(i).Name, .VBComponents(i).Type)
        Next
    End With
  
    '�o�͐�Ƃ��Ẵ��[�N�V�[�g����������
    Dim sheetName As String: sheetName = "ExportedModules"
    Dim r As Boolean
    r = KzCreateWorksheetInWorkbook(wb, sheetName)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    
    '�v���V�W���̏����V�[�g�ɏo�͂���
    Dim v
    With ws
        .Cells.Clear
        .Range("A1:I1").Value = Array("Project", "ModuleType", "Module", "Scope", "ProcKind", "Procedure", "LineNo", "Source", "Comment")
        .Range("A1:H1").Interior.Color = RGB(200, 200, 200) ' �w�i�F���O���[
        i = 2
        For Each v In dicProcInfo.Items
            .Cells(i, 1) = wb.VBProject.Name   ' KazurayamVbaLib
            .Cells(i, 2) = v.ModuleType        ' Standard | Class (Sheet Module, ThisWorkbook Module, Userforms�͖��T�|�[�g)
            .Cells(i, 3) = v.ModName           ' KzSensible
            .Cells(i, 4) = v.Scope             ' Public | Private | Static
            ' .Cells(i, 5) = v.ProcKindName
            .Cells(i, 5) = FormatProcKindName(v.procKindName, v.Source)
            .Cells(i, 6) = v.ProcName          ' KzProcedureList
            .Cells(i, 7) = v.LineNo
            .Cells(i, 8) = v.Source
            .Cells(i, 9) = "'" & v.Comment
            i = i + 1
        Next
        Cells.EntireRow.AutoFit
        Cells.EntireColumn.AutoFit
        Range("F1").ColumnWidth = 30
        Range("H1:I1").ColumnWidth = 100
        
    End With

    '�V�[�g�̍s���v���W�F�N�g��>���W���[����>�v���V�W�����̏����Ń\�[�g����
    With ws.Sort
        With .SortFields
            .Clear
            .Add key:=ws.Range("A2"), Order:=xlAscending
            .Add key:=ws.Range("C2"), Order:=xlAscending
            .Add key:=ws.Range("F2"), Order:=xlAscending
        End With
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(i, 9))
        .Header = xlYes
        .Apply
    End With
    
    '�s�̍������������߂���
    ws.Rows.AutoFit
    
    Set dicProcInfo = Nothing
End Sub



Private Function FormatProcKindName(ByVal procKindName As String, ByVal Source As String) As String
    If (InStr(1, LCase(Source), " function ") > 0) Then
        FormatProcKindName = "Function"
    ElseIf (InStr(1, LCase(Source), " sub ") > 0) Then
        FormatProcKindName = "Sub"
    Else
        FormatProcKindName = procKindName  ' Property Let | Property Set | Property Get
    End If
End Function



