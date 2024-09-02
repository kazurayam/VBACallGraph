Attribute VB_Name = "BbUtil"
Option Explicit

'BbUtil


Public Function VarTypeAsString(ByVal var As Variant) As String
    ' ����var��type�𒲂ׂĕϐ��̌^������������i"Integer"�Ȃǁj��Ԃ�
    Dim typeValue As Long: typeValue = VarType(var)
    Dim result As String: result = "unknown"
    If typeValue = 2 Then
        result = "Integer"
    ElseIf typeValue = 3 Then
        result = "Long"
    ElseIf typeValue = 5 Then
        result = "Double"
    ElseIf typeValue = 8 Then
        result = "String"
    ElseIf typeValue = 11 Then
        result = "Boolean"
    ElseIf typeValue = 7 Then
        result = "Date"
    ElseIf typeValue = 9 Then
        result = "Object"
    ElseIf typeValue = 0 Then
        result = "Variant"
    ElseIf typeValue = 8200 Then
        result = "String()"
    ElseIf typeValue = 8194 Then
        result = "Integer()"
    Else
        result = Str(typeValue)
    End If
    VarTypeAsString = result
End Function


Public Function ResolveExternalFilePath( _
        ByVal theWorkbook As Workbook, _
        ByVal sheetName As String, _
        ByVal rangeLiteral As String) As String
    'theWorkbook�Ƃ��ė^����ꂽ���[�N�u�b�N�̂Ȃ���
    'sheetName�Ƃ��ė^����ꂽ���[�N�V�[�g�������āA���̒���
    'rangeLiteral�Ƃ��ė^����ꂽ�Z���������āA���̂Ȃ���
    '�O���t�@�C���̃p�X�������Ă���Ɗ��҂���B
    '���̃p�X��theWorkbook�����Ƃ��鑊�΃p�X�ł���Ɗ��҂���B
    '�O���t�@�C���̃p�X�𔭌����A������΃p�X�ɕϊ����āAFunction�̒l�Ƃ��ĕԂ��B
    '���̊֐���.xlsm�t�@�C���̉��������߂�̂ɗL�p�ł���B
    '.xlsm�t�@�C�����猩���O���t�@�C���̃p�X��VBA�R�[�h�̂Ȃ���
    '�Œ�l�Ƃ��ď����̂ł͂Ȃ��A
    '���[�N�V�[�g�̃Z���̒l�Ƃ��ď������Ƃ��\�ɂ���B
    If BbWorksheet.IsWorksheetPresentInWorkbook(theWorkbook, sheetName) Then
        Dim ws As Worksheet: Set ws = theWorkbook.Worksheets(sheetName)
        Dim path As String
        path = ws.Range(rangeLiteral)
        ResolveExternalFilePath = BbFile.AbsolutifyPath(BbFile.ToLocalFilePath(theWorkbook.path), path)
    Else
        Debug.Print theWorkbook.FullName + " does not have a worksheet named " + sheetName
        ResolveExternalFilePath = ""
    End If
End Function


