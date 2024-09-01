Attribute VB_Name = "WriteSettlement"

Option Explicit

'AccountSum���W���[��

'�����o�[�����[�N�u�b�N�̃f�[�^����͂Ƃ��āy�����o�[�L�^�z�V�[�g���쐬����B
'�y�����o�[�L�^�z�V�[�g�� ����Ȗڂ��Ƃ̏��v Sum of account ���Z�o����B
'����Ȗڂ��Ƃ̏��v��₢���킹�邱�Ƃ̂ł��� Get���v Function��񋟂���B
'���Z���V�[�g�Ȃǂق��̃V�[�g�̃Z�����Q�Ǝ��� Get���v �֐����g����
'����Ȗڂ��Ƃ̏��v���Q�Ƃ��邱�Ƃ��ł���B

Const TheSheetName As String = "�����o�[�L�^"
Const RecTableName As String = "�e�[�u�������o�[�L�^"
Const SumTableName As String = "�e�[�u������Ȗڂ��Ƃ̏��v"

Enum enumSource
    �̎����ԍ� = 1
    �ߘa
    ��
    ��
    �����Ȗ�
    �����⏕�Ȗ�
    �x�o�Ȗ�
    �x�o�⏕�Ȗ�
    ���x�񍐒P��
    �E�v
    �ؕ����z
    �ݕ����z
    �����c��
    ���֎�
    ���Z�ς݂�
End Enum

Enum enumTarget
    ����Ȗ� = 1
    �N����
    �����Ȗ�
    �����⏕�Ȗ�
    �x�o�Ȗ�
    �x�o�⏕�Ȗ�
    ���x�񍐒P��
    �E�v
    �ؕ����z
    �ݕ����z
End Enum


'==============================================================================

Public Sub ���o���L�^����荞��(ByVal wsCashSource As Worksheet, _
                                ByVal wsCashTarget As Worksheet, _
                                Optional ByVal cashbookTableName As String = "CashbookTable1", _
                                Optional ByVal periodStart As Date = #4/1/2022#, _
                                Optional ByVal periodEnd As Date = #3/31/2023#, _
                                Optional ofReportingUnit = "*", _
                                Optional positiveLike = True, _
                                Optional initializeTarget = True)
    Dim tblSource As ListObject
    Set tblSource = wsCashSource.ListObjects(cashbookTableName)
    Dim tblTarget As ListObject
    Set tblTarget = wsCashTarget.ListObjects(RecTableName)
     
    'initializeTarget�I�v�V������True�Ȃ�΁@�^�[�Q�b�g�̃e�[�u��������������
    If initializeTarget Then
        tblTarget.DataBodyRange.Delete
    End If
    
    '�\�[�X�̃e�[�u������^�[�Q�b�g�̃e�[�u���ɍs��]�ʂ���B
    '�y�ߘa�z+�y���z+�y���z���y�N�����z�ɕϊ�������
    '�s�v�ȗ�𖳎������肵�Ȃ���B
    Dim i As Long
    For i = 1 To tblSource.ListRows.Count
        Dim rowSource As ListRow: Set rowSource = tblSource.ListRows(i)
        If Is��荞�ނׂ�(rowSource, periodStart, periodEnd, ofReportingUnit, positiveLike) Then
            '�f�[�^���ʂ�
            Call TransferRow(rowSource, tblTarget)
        End If
    Next i
    'A��i����Ȗځj�̕����w�肷��
    wsCashTarget.Columns("A").ColumnWidth = 22
    'B��i�N�����j�̕���K���ɂ���
    wsCashTarget.Columns("B:B").AutoFit
End Sub


'rowSource�f�[�^�P�s��tblTarget�W�I�ł���e�[�u����Add����
Public Sub TransferRow(rowSource As ListRow, tblTarget As ListObject)
    With tblTarget.ListRows.Add
        .Range(enumTarget.�N����).value = ToDate( _
                          rowSource.Range(enumSource.�ߘa).value, _
                            rowSource.Range(enumSource.��).value, _
                            rowSource.Range(enumSource.��).value)
        .Range(enumTarget.�����Ȗ�).value = rowSource.Range(enumSource.�����Ȗ�).value
        .Range(enumTarget.�����⏕�Ȗ�).value = rowSource.Range(enumSource.�����⏕�Ȗ�).value
        .Range(enumTarget.�x�o�Ȗ�).value = rowSource.Range(enumSource.�x�o�Ȗ�).value
        .Range(enumTarget.�x�o�⏕�Ȗ�).value = rowSource.Range(enumSource.�x�o�⏕�Ȗ�).value
        .Range(enumTarget.���x�񍐒P��).value = rowSource.Range(enumSource.���x�񍐒P��).value
        .Range(enumTarget.����Ȗ�).value = To����Ȗ�(rowSource)
        .Range(enumTarget.�E�v).value = rowSource.Range(enumSource.�E�v).value
        .Range(enumTarget.�ؕ����z).value = rowSource.Range(enumSource.�ؕ����z).value
        .Range(enumTarget.�ݕ����z).value = rowSource.Range(enumSource.�ݕ����z).value
        
        '���z�̃Z���̏�����ݒ肷��i�ʉ݁A3�����J���}�ŋ�؂�A\�����j
        .Range(enumTarget.�ؕ����z).NumberFormatLocal = "#,###"
        .Range(enumTarget.�ݕ����z).NumberFormatLocal = "#,###"
    End With
End Sub


Public Function Is��荞�ނׂ�(ByVal rowSource As ListRow, _
                                Optional ByVal periodStart As Date = #4/1/2022#, _
                                Optional ByVal periodEnd As Date = #3/31/2023#, _
                                Optional ofReportingUnit = "*", _
                                Optional positiveLike = True) As Boolean
    If Is���z����[����(rowSource) Then
        If Is�N�x����(rowSource, periodStart, periodEnd) Then
            If Is���x�񍐒P�ʂ��Y������(rowSource, ofReportingUnit, positiveLike) Then
                Is��荞�ނׂ� = True
            Else
                Is��荞�ނׂ� = False
            End If
        Else
            Is��荞�ނׂ� = False
        End If
    Else
        Is��荞�ނׂ� = False
    End If
End Function


Public Function Is���z����[����(ByVal rowSource As ListRow) As Boolean
    If rowSource.Range(enumSource.�ݕ����z).value <> 0 Or _
        rowSource.Range(enumSource.�ؕ����z).value <> 0 Then
        Is���z����[���� = True
    Else
        Is���z����[���� = False
    End If
End Function


Public Function Is�N�x����(ByVal rowSource As ListRow, _
                            Optional ByVal periodStart As Date = #4/1/2022#, _
                            Optional ByVal periodEnd As Date = #3/31/2023#) As Boolean
    Dim dt As Date
    dt = ToDate(rowSource.Range(enumSource.�ߘa).value, _
                rowSource.Range(enumSource.��).value, _
                rowSource.Range(enumSource.��).value)
    If periodStart <= dt And dt <= periodEnd Then
        Is�N�x���� = True
    Else
        Is�N�x���� = False
    End If
End Function


Public Function Is���x�񍐒P�ʂ��Y������(ByVal rowSource As ListRow, _
                                         Optional ofReportingUnit = "*", _
                                         Optional positiveLike = True) As Boolean
    Dim rpUnit As String: rpUnit = rowSource.Range(enumSource.���x�񍐒P��)
    If (positiveLike And rpUnit Like ofReportingUnit) Or _
        (Not positiveLike And Not rpUnit Like ofReportingUnit) Then
        Is���x�񍐒P�ʂ��Y������ = True
    Else
        Is���x�񍐒P�ʂ��Y������ = False
    End If
End Function


'==============================================================================


'���͍s�́y�����Ȗځz�y�����⏕�Ȗځz�y�x�o�Ȗځz�y�x�o�⏕�Ȗځz����
'  "������/�ʐM��"
'  "���/A���"
'�̂悤�ȕ�������������ĕԂ��B������w���o���L�^�x�V�[�g�́y����Ȗځz�ɗ��p����B
'�y����Ȗځz�Ɓy�N�����z���L�[�Ƃ��ăe�[�u���̍s���\�[�g���邽�߂ɁB
'VBA�̃e�[�u����Sort�@�\�̓L�[��3�܂ł����w��ł��Ȃ�����A�y����Ȗځz���������邱�Ƃ��K�v������
Public Function To����Ȗ�(ByVal rowSource As ListRow) As String
    If rowSource.Range(enumSource.�x�o�Ȗ�).value <> "" And _
        rowSource.Range(enumSource.�x�o�⏕�Ȗ�).value <> "" Then
        
        To����Ȗ� = "�x�o" & "/" _
                        & rowSource.Range(enumSource.�x�o�Ȗ�).value _
                        & "/" & rowSource.Range(enumSource.�x�o�⏕�Ȗ�).value
    
    ElseIf rowSource.Range(enumSource.�����Ȗ�).value <> "" And _
            rowSource.Range(enumSource.�����⏕�Ȗ�).value <> "" Then
            
        To����Ȗ� = "����" & "/" _
                        & rowSource.Range(enumSource.�����Ȗ�).value _
                        & "/" & rowSource.Range(enumSource.�����⏕�Ȗ�).value
    Else
        To����Ȗ� = "?/?"
    End If

End Function


Public Function ToDate(ByVal YY As Long, ByVal MM As Long, ByVal DD As Long) As Date
    ' �NYY�ƌ�MM�Ɠ�DD������t�𐶐��������Date�Ƃ��ĕԂ�
    ' �����o�[���̔NYY�͘a��ł���͂�������A�K�؂ɕϊ�����B
    ' �ߘa�𐼗�N�ɕϊ�����ɂ́i�蔲�����Ƃ���jYY�ɐ���2018��+���邱�Ƃɂ���
    Dim d As Date
    d = DateSerial(YY + 2018, MM, DD)
    ToDate = d
End Function


Public Sub ���o���L�^���\�[�g����(ByVal wsCashTarget As Worksheet)
    Dim tblTarget As ListObject
    Set tblTarget = wsCashTarget.ListObjects(RecTableName)
    With tblTarget
        .Range.Sort key1:=.ListColumns(enumTarget.����Ȗ�), order1:=xlAscending, _
                    key2:=.ListColumns(enumTarget.�N����), order2:=xlAscending, _
                    Header:=xlYes
    End With
End Sub


'�w�����o�[�L�^�x���[�N�V�[�g��[�e�[�u������Ȗڂ��Ƃ̏��v]������B
'�����Ƀf�[�^�𖄂߂�B
'�w�����o�[�L�^�x���[�N�V�[�g��[�e�[�u�������o�[�L�^]�����邩��A���̍s��
'����Ȗڂ��Ƃ̕ʃV�[�g�ɏ��������čs���R�s�[���A�e�[�u���ɕϊ����ďW�v���Z�o����B
'�����Ċ���ȖڂƏ��v�������z��[�e�[�u�������o�[�L�^]�ɏ������ށB
'����Ȗڂ��Ƃ̕ʃV�[�g�͏W�v���ς߂Ζ��p�Ȃ̂ō폜����B

Public Sub ���v�̕\�����(ByRef ws�����o�[�L�^ As Worksheet)
    Dim tbl�L�^ As ListObject
    Set tbl�L�^ = ws�����o�[�L�^.ListObjects(RecTableName)
    Dim tbl���v As ListObject
    Set tbl���v = ws�����o�[�L�^.ListObjects(SumTableName)
    
    '[�e�[�u������Ȗڂ��Ƃ̏��v]�̃f�[�^�s�����������������
    If tbl���v.ListRows.Count <> 0 Then
        tbl���v.DataBodyRange.Delete
    End If
    
    '����Ȗڂ̈ꗗ�i�d���Ȃ��j���擾����
    Dim unique����Ȗږ��̗� As Variant
    unique����Ȗږ��̗� = KzRange.KzGetUniqueItems(tbl�L�^.ListColumns(1).DataBodyRange)
    
    '����Ȗڂ��Ƃ̏��v���Z�o����[�e�[�u������Ȗڂ��Ƃ̏��v]�ɍs�Ƃ��đ}������
    Dim column As Variant
    For Each column In unique����Ȗږ��̗�
        Call ����Ȗڂ̏��v���Z�o����(ws�����o�[�L�^, column)
    Next

End Sub


Public Sub Test���v�̕\�����()
    Call BbLog.Clear
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(TheSheetName)
    Call ���v�̕\�����(ws)
End Sub


'==============================================================================

Public Sub ����Ȗڂ̏��v���Z�o����(ByRef ws As Worksheet, ByVal ����Ȗږ� As String)
    'Debug.Print "���v���Z�o����: " & ����Ȗږ�
    'Debug.Print "enumTarget.����Ȗ� = " & enumTarget.����Ȗ�
    
    Dim tbl�L�^ As ListObject
    Set tbl�L�^ = ws.ListObjects(RecTableName)
    Dim tbl���v As ListObject
    Set tbl���v = ws.ListObjects(SumTableName)
    
    '����Ȗږ����L�[�Ƃ���t�B���^��[�e�[�u�������o�[�L�^]�ɓK�p����B
    '�w�肳�ꂽ����ȖڂɊY������s��I�ʂ��Ȃ�����o���f�[�^�̍s�𑖍�����B
    '�ݕ����z�̍��v�Ǝؕ����z�̍��v���Z�o����B
    '�S���̍s�𒲂׏I�������[�e�[�u������Ȗڂ��Ƃ̏��v]�ɍs��Add����
    '����ȖڂƑݕ����z�̍��v�Ǝؕ����z�̍��v�������o���B
    
    Dim debt As Long: debt = 0  '�ؕ��i���j
    Dim credit As Long: credit = 0  '�ݕ��i�E�j
    Dim acc As String
    
    With tbl�L�^
        '�t�B���^��K�p����
        .Range.AutoFilter 1, ����Ȗږ�
        '�f�[�^�s����Ԃ𑖍�����
        Dim i As Long
        For i = 1 To .ListRows.Count
            '�t�B���^�ɂ���ĕ����яオ�����s��I������
            If .ListRows(i).Range.EntireRow.Hidden = False Then
                '���v���Z�o����
                acc = .ListRows(i).Range(enumTarget.����Ȗ�).value
                If acc Like "����/*" Then
                    '����Ȗڂ������Ȃ�
                    debt = debt + .ListRows(i).Range(enumTarget.�ؕ����z) - .ListRows(i).Range(enumTarget.�ݕ����z)
                ElseIf acc Like "�x�o/*" Then
                    '����Ȗڂ��x�o�Ȃ�
                    credit = credit + .ListRows(i).Range(enumTarget.�ݕ����z) - .ListRows(i).Range(enumTarget.�ؕ����z)
                End If
                'Debug.Print acc & " " & debt & " " & credit
            End If
        Next
        '�t�B���^������
        .Range.AutoFilter enumTarget.����Ȗ�
    End With
    
    '[�e�[�u������Ȗڂ��Ƃ̏��v]�ɍs��}�����ĎZ�o�������v����������
    With tbl���v.ListRows.Add
        .Range(enumTarget.����Ȗ�).value = acc
        .Range(enumTarget.�ؕ����z).value = debt
        .Range(enumTarget.�ݕ����z).value = credit
    
        '���z�̃Z���̏�����ݒ肷��i�ʉ݁A3�����J���}�ŋ�؂�A\�����j
        .Range(enumTarget.�ؕ����z).NumberFormatLocal = "#,###"
        .Range(enumTarget.�ݕ����z).NumberFormatLocal = "#,###"
    End With
End Sub

' ����Ȗڂ̏��v���Z�o����Sub���e�X�g����
Public Sub Test����Ȗڂ̏��v���Z�o����()
    Call BbLog.Clear
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TheSheetName)
    '[�e�[�u������Ȗڂ��Ƃ̏��v]�̃f�[�^�s�����������������
    Dim tbl���v As ListObject
    Set tbl���v = ws.ListObjects(SumTableName)
    If tbl���v.ListRows.Count <> 0 Then
        tbl���v.DataBodyRange.Delete
    End If
    Call ����Ȗڂ̏��v���Z�o����(ws, "�x�o/�c����/�c����")
    Call ����Ȗڂ̏��v���Z�o����(ws, "�x�o/������/�U���萔��")
    Call ����Ȗڂ̏��v���Z�o����(ws, "����/���/A���")
End Sub


'==============================================================================


Public Sub �����o�[�L�^���[�N�V�[�g��������΍��(ByRef Workbook As Workbook, ByVal sheetName As String)
    'ThisWorkbook�̂Ȃ��Ɂy�����o�[�L�^�z���[�N�V�[�g�����łɂ��邩�ǂ����𒲂ׂ�B
    '���������łɂ�������Ȃɂ������ɂ����܂��B
    '�܂�����������y�����o�[�L�^�z���[�N�V�[�g�����A"CashbookProj"�̂Ȃ���
    '�p�ӂ��ꂽ�y�����o�[�L�^�e���v���[�g�z���[�N�V�[�g�̓��e���R�s�[���ĕۑ�����B
    If Not BbWorksheet.IsWorksheetPresentInWorkbook(Workbook, sheetName) Then
        Call BbLog.Info("AccountSum", "�����o�[�L�^���[�N�V�[�g��������΍��", sheetName & "���[�N�V�[�g�����������̂ō��܂�")
        Dim wsNew As Worksheet: Set wsNew = Worksheets.Add
        wsNew.Name = sheetName
        wsNew.Activate
        
        '[�e�[�u������Ȗڂ��Ƃ̏��v]�����
        Range("A2").value = "����Ȗ�"
        Range("B2").value = "��2"
        Range("C2").value = "��3"
        Range("D2").value = "��4"
        Range("E2").value = "��5"
        Range("F2").value = "��6"
        Range("G2").value = "��7"
        Range("H2").value = "��8"
        Range("I2").value = "�ؕ����z"
        Range("J2").value = "�ݕ����z"
        Range("A3").value = "?"
        Range("I3").value = 0
        Range("J3").value = 0
        With wsNew
            .ListObjects.Add 1, Range("A3").CurrentRegion
            .ListObjects(1).Name = SumTableName
            .ListObjects(1).TableStyle = "TableStyleLight8"
            .ListObjects(1).ShowTotals = True
        End With
        Range("A1").value = "����Ȗڂ��Ƃ̏��v"
        Range("A1").Style = "���o�� 2"
        
        
        '[�e�[�u�������o�[�L�^]�����
        Range("A9").value = "����Ȗ�"
        Range("B9").value = "�N����"
        Range("C9").value = "�����Ȗ�"
        Range("D9").value = "�����⏕�Ȗ�"
        Range("E9").value = "�x�o�Ȗ�"
        Range("F9").value = "�x�o�⏕�Ȗ�"
        Range("G9").value = "���x�񍐒P��"
        Range("H9").value = "�E�v"
        Range("I9").value = "�ؕ����z"
        Range("J9").value = "�ݕ����z"
        Range("A10").value = "?"
        Range("B10").value = ""
        Range("C10").value = ""
        Range("D10").value = ""
        Range("E10").value = ""
        Range("F10").value = ""
        Range("G10").value = ""
        Range("H10").value = ""
        Range("I10").value = 0
        Range("J10").value = 0
        With wsNew
            .ListObjects.Add 1, Range("A10").CurrentRegion
            .ListObjects(2).Name = RecTableName
            .ListObjects(2).TableStyle = "TableStyleLight9"
            .ListObjects(2).ShowTotals = True
        End With
        Range("A8").value = "�����o�[�L�^"
        Range("A8").Style = "���o�� 2"
        
        '�E�C���h�E�g���Œ肷��@C3�Z����
        Range("C3").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        
        
    End If
    
End Sub

'==============================================================================

'�y�����o�[�L�^�z���[�N�V�[�g��[�e�[�u������Ȗږ��̏��v]����������
'�w�肳�ꂽ����Ȗڂ̏��v���z��Ԃ��B
'��@����Ȗږ��u�x�o/��c��/�������v�ɂ������� 33000 ���Ԃ����
Public Function Get���v(ByVal ����Ȗږ� As String) As Long
    Dim ws As Worksheet: Set ws = Worksheets(TheSheetName)
    Dim tbl���v As ListObject: Set tbl���v = ws.ListObjects(SumTableName)
    Dim val As Long: val = 0
    With tbl���v
        Dim i As Long
        For i = 1 To .ListRows.Count
            Dim acc As String: acc = .ListRows(i).Range(enumTarget.����Ȗ�).value
            If acc Like ����Ȗږ� Then
                If acc Like "����/*" Then
                    val = .ListRows(i).Range(enumTarget.�ؕ����z)
                ElseIf acc Like "�x�o/*" Then
                    val = .ListRows(i).Range(enumTarget.�ݕ����z)
                Else
                    val = -1
                End If
                Exit For
            End If
        Next
    End With
    Get���v = val
End Function


Public Sub Test_Get���v()
    Call BbLog.Clear
    Dim val As Long
    val = Get���v("�x�o/�c����/�c����")
    Debug.Assert val > 0
End Sub





