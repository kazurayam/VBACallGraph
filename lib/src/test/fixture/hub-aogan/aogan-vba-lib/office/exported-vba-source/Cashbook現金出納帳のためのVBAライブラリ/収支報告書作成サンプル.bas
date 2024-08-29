Attribute VB_Name = "���x�񍐏��쐬�T���v��"
Option Explicit

' ���[�N�V�[�g�uwork�v�̂Ȃ��́u��S�T��T�u���b�N�u�K���v�񍐁v�̂Ȃ��ɐ������������ށB
' �{�v���W�F�N�g��src�t�H���_�̉��Ɂu�����o�[���X����Ȉ��ߘa4�N�x.xlsm�v������B
' ���̃��[�N�u�b�N�́u�����o�[���v���[�N�V�[�g�ɏ����ꂽ���o���f�[�^����͂Ƃ��ĎQ�Ƃ���B

Sub Main_TohokuBlockLectures45th()

    Call KzCls
    Debug.Print "Main_TohokuBlockLectuers45th STARTED"
    
    ThisWorkbook.Activate
    
    '�o�͐�ƂȂ郏�[�N�V�[�g����肷��
    Dim outWs As Worksheet
    Set outWs = ActiveWorkbook.Worksheets("work")
    
    '���ׂ̏o�͐�ƂȂ�e�[�u����錾����
    Dim fullTbl As ListObject: Set fullTbl = outWs.ListObjects("�e�[�u��1")
    '���ׂ̃e�[�u��������������
    '�e�[�u������łȂ����Ƃ��m�F���Ă���Delete����
    If Not fullTbl.DataBodyRange Is Nothing Then
        fullTbl.DataBodyRange.Delete
    End If
    
    '���x�񍐒P�ʂ�錾����
    Dim rpUnit As String: rpUnit = "���k�u���b�N�u�K��"

    '===================================================================================
    '���k��Ȉ��A����X����\�̌����o�[���ł���Excel�t�@�C�����J��
    ThisWorkbook.Activate
    Debug.Print ">>" & KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B3")
    Dim wb1 As Workbook
    Set wb1 = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B3"))
    Dim cb1 As Cashbook: Set cb1 = CreateCashbook(wb1, "�����o�[��", "CashbookTable2")
    Call PrintAccounts(cb1, rpUnit)
    '
    Dim cs1 As CashSelector: Set cs1 = New CashSelector
    Call cs1.Initialize(cb1, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    
    '�W�v
    Call TranscribeSum(cs1, rpUnit, AccountType.Income, "�G����", "�Z�~�i�[�Q����", outWs.Range("$G$7"))
    Call TranscribeSum(cs1, rpUnit, AccountType.Income, "�G����", "�L��������", outWs.Range("$G$8"))
    
    '����
    Call TranscribeDetail(cs1, rpUnit, AccountType.Income, "�G����", "�Z�~�i�[�Q����", fullTbl)
    Call TranscribeDetail(cs1, rpUnit, AccountType.Income, "�G����", "�L��������", fullTbl)
    
    '��n���B�O����Excel�t�@�C�������
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb1.Close
    Set cs1 = Nothing
    Set cb1 = Nothing
    Set wb1 = Nothing
    
    '===================================================================================
    '�X����Ȉ��̌����o�[���ł���Excel�t�@�C�����J��
    ThisWorkbook.Activate
    Debug.Print ">>" & KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2")
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Dim cb2 As Cashbook: Set cb2 = CreateCashbook(wb2, "�����o�[��", "CashbookTable1")
    '�X����Ȉ��̌����o�[���ɏ����ꂽ���o���f�[�^�Łu���k�u���b�N�u�K��v�Ɋ֘A���銨��Ȗڂ�񋓂���
    Call PrintAccounts(cb2, rpUnit)
    
    'CashSelector�I�u�W�F�N�g�𐶐����ACashbook�I�u�W�F�N�g���Q�Ƃ���悤�ݒ肷��
    Dim cs2 As CashSelector: Set cs2 = New CashSelector
    Call cs2.Initialize(cb2, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    
    '�W�v
    '����Ȗڂ��Ƃ̖��ׂ�₢���킹�A���z�̏��v�����߂�B
    Call TranscribeSum(cs2, rpUnit, AccountType.Expense, "���Ɣ�", "�w�p��", outWs.Range("$F$10"))
    Call TranscribeSum(cs2, rpUnit, AccountType.Expense, "���Ɣ�", "�ʐM��", outWs.Range("$F$11"))
    Call TranscribeSum(cs2, rpUnit, AccountType.Expense, "������", "�ʐM��", outWs.Range("$F$12"))
    
    '����
    Call TranscribeDetail(cs2, rpUnit, AccountType.Expense, "���Ɣ�", "�w�p��", fullTbl)
    Call TranscribeDetail(cs2, rpUnit, AccountType.Expense, "���Ɣ�", "�ʐM��", fullTbl)
    Call TranscribeDetail(cs2, rpUnit, AccountType.Expense, "������", "�ʐM��", fullTbl)
    
    '��n���B�O����Excel�t�@�C�������
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb2.Close
    Set cs2 = Nothing
    Set cb2 = Nothing
    Set wb2 = Nothing
    
    
    Debug.Print "Main_TohokuBlockLectuers45th FINISHED"
End Sub


Private Sub TranscribeSum(ByRef cs As CashSelector, _
                                ByVal theReportingUnit As String, _
                                ByVal AccType As AccountType, _
                                ByVal accName As String, _
                                ByVal subAccName As String, _
                                ByRef targetCell As Range)
    '�w�肳�ꂽ����Ȗڂ̓��o���̋��z�����Z���A�w�肳�ꂽ�Z���ɓ]�L����
    
    'Debug.Print theReportingUnit & "," & accType & "," & accName & "," & subAccName & "," & targetCell
    
    '�p�����[�^�Ŏw�肳�ꂽ
    Dim selected As CashList
    If AccType = AccountType.Expense Then
        Set selected = cs.SelectCashList(AccType, accName, subAccName, theReportingUnit)
        targetCell.value = selected.SumOfExpenseAmount()
    Else
        Set selected = cs.SelectCashList(AccType, accName, subAccName, theReportingUnit)
        targetCell.value = selected.SumOfIncomeAmount()
    End If
    'tearDown
    Set selected = Nothing
End Sub

Private Sub PrintAccounts(ByVal cb As Cashbook, ByVal rpUnit As String)
    Dim accFinder As AccountsFinder: Set accFinder = New AccountsFinder
    Call accFinder.Initialize(cb, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    Dim dic As Dictionary: Set dic = accFinder.FindAccounts(rpUnit)
    Dim key As Variant
    For Each key In dic
        Debug.Print key & ":" & dic(key)
    Next
End Sub


Private Sub TranscribeDetail(ByVal cs As CashSelector, _
                            ByVal rpUnit As String, _
                            ByVal AccType As AccountType, _
                            ByVal ofAccountName As String, _
                            ByVal ofSubAccountName As String, _
                            ByRef targetTable As ListObject)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccType, ofAccountName, ofSubAccountName, rpUnit)
    Dim i As Long
    Dim ch As Cash
    For i = 1 To selected.Count
        Set ch = selected.Items(i)
        With targetTable.ListRows.Add
            .Range(1).value = ch.ItsAccount.AccountTypeAsString
            .Range(2).value = ch.ItsAccount.AccountName
            .Range(3).value = ch.ItsAccount.SubAccountName
            .Range(4).value = ch.Description
            If ch.ExpenseAmount > 0 Then
                .Range(5).value = ch.ExpenseAmount
            End If
            If ch.IncomeAmount > 0 Then
                .Range(6).value = ch.IncomeAmount
            End If
            .Range(7).value = ch.ToDate()
        End With
    Next i
End Sub
