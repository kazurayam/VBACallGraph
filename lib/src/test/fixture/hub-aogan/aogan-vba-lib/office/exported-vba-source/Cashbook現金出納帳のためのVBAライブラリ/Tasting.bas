Attribute VB_Name = "Tasting"
                                                                                                                                                                                                                                                Option Explicit

' Tasting : ��������A�C�y�ɐG���Ă݂�A������Ǝ����A���낢���

Sub TasteCashbook()
    ' Cashbook�N���X�𖡌�����
    ' Cashbook�I�u�W�F�N�g�𐶐����āACount�v���p�e�B��ǂ݂�����Print����
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    'Act:
    Dim cb As Cashbook: Set cb = CreateCashbook(wb)
    'Assert:
    Debug.Print "cb.Count=" & cb.Count
    'TearDown
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
End Sub


Sub TasteCash_ColumnHeader()
    ' Cash�I�u�W�F�N�g��ColumnHeader()���\�b�h�𖡌�����
    Dim cs As Cash
    Set cs = New Cash
    Debug.Print cs.ColumnHeader()
End Sub


Sub TasteCash_ToString()
    ' Cash�I�u�W�F�N�g��ToString()���\�b�h�𖡌�����
    Call KzCls
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("�����o�[��")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    '
    Dim data As Range: Set data = tbl.ListRows(12).Range '�����̈��
    Dim cs As Cash: Set cs = New Cash
    Call cs.Initialize(data)
    Debug.Print cs.ToString()
    '
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
End Sub


Sub TasteCash_ToDate()
    ' Cash�I�u�W�F�N�g��ToDate()���\�b�h�𖡌�����
    Call KzCls
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("�����o�[��")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    '
    Dim data As Range: Set data = tbl.ListRows(12).Range '�����̈��
    Dim cs As Cash: Set cs = New Cash
    Call cs.Initialize(data)
    Debug.Print cs.ToDate()
    'TearDown
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
End Sub

Sub TasteCashSelector_SelectCashList_ofIncome()
    ' CashSelector�I�u�W�F�N�g��SelectCashList���\�b�h�𖡌�����
    ' Income�̏ꍇ
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("�����o�[��")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    '
    Call KzCls
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "�G����", "�Z�~�i�[�Q����", "��ȃt�H�[����")
    Debug.Print selected.Count
    Debug.Print selected.ToString()
    'TearDown
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
End Sub


Sub TasteCashSelector_SelectCashList_ofExpense()
    ' CashSelector�I�u�W�F�N�g��SelectCashListOfIncome���\�b�h�𖡌�����
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("�����o�[��")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    '
    Call KzCls
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Expense, "���Ɣ�", "�L���", "���")
    
    Debug.Print selected.Count
    Debug.Print selected.ToString()
    'TearDown
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
End Sub


Sub TasteAccount()
    Call KzCls
    ' Expense account
    Dim accExpense As Account: Set accExpense = New Account
    Call accExpense.Initialize(AccountType.Expense, "���Ɣ�", "�L���")
    Debug.Print "accExpense.accType: " & accExpense.AccType
    Debug.Print "accExpense.AccountName: " & accExpense.AccountName
    Debug.Print "accExpense.SubAccountName: " & accExpense.SubAccountName
    Debug.Print "accExpense.ToString(): " & accExpense.ToString()
    ' Income account
    Dim accIncome As Account: Set accIncome = New Account
    Call accIncome.Initialize(AccountType.Income, "�G����", "�Z�~�i�[�Q����")
    Debug.Print "accIncome.accType: " & accIncome.AccType
    Debug.Print "accIncome.AccountName: " & accIncome.AccountName
    Debug.Print "accIncome.SubAccountName: " & accIncome.SubAccountName
    Debug.Print "accIncome.ToString(): " & accIncome.ToString()
End Sub


Sub TasteAccountsFinder_allUnit()
    ' AccountsFinder�N���X�𖡌�����
    'Arrange
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("�����o�[��")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act
    Dim accFinder As AccountsFinder: Set accFinder = New AccountsFinder
    Call accFinder.Initialize(cb)
    Dim dic As Dictionary
    Set dic = accFinder.FindAccounts()
    'Assert
    Call KzCls
    Dim key As Variant
    For Each key In dic
        Debug.Print key & ":" & dic(key)
    Next
'0//:1
'����/���/A���:44
'����/���/B���:32
'����/���/C���:1
'����/�⏕��/�⏕��:1
'����/�G����/�Z�~�i�[�Q����:4
'����/�G����/�L��������:12
'����/�G����/�G����:1
'�x�o/������/���Օi��:27
'�x�o/������/�ʐM��:58
'�x�o/���Ɣ�/���O�q����:1
'�x�o/���Ɣ�/�w�p��:17
'�x�o/���Ɣ�/�L���:3
'�x�o/���Ɣ�/������:1
'�x�o/���Ɣ�/��ȃR�E���f�B�J���֌W��:1
'�x�o/���Ɣ�/�ʐM��:1
'�x�o/��c��/�������:1
'�x�o/�o����/�o������⏕:1
'�x�o/�c����/�c����:1

    Debug.Print vbNewLine & "----------- by FindKeysAsString"
    Debug.Print accFinder.FindKeysAsString()
'0//
'����/���/A���
'����/���/B���
'����/���/C���
'����/�⏕��/�⏕��
'����/�G����/�Z�~�i�[�Q����
'����/�G����/�L��������
'����/�G����/�G����
'�x�o/������/���Օi��
'�x�o/������/�ʐM��
'�x�o/���Ɣ�/���O�q����
'�x�o/���Ɣ�/�w�p��
'�x�o/���Ɣ�/�L���
'�x�o/���Ɣ�/������
'�x�o/���Ɣ�/��ȃR�E���f�B�J���֌W��
'�x�o/���Ɣ�/�ʐM��
'�x�o/��c��/�������
'�x�o/�o����/�o������⏕
'�x�o/�c����/�c����
'
    'TearDown
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
End Sub



Sub TasteAccountsFinder_specifyReportingUnit()
    ' AccountsFinder�N���X�𖡌����� ���������x�񍐒P�ʂ���肵��
    'Arrange
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Dim ws As Worksheet: Set ws = wb.Worksheets("�����o�[��")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("CashbookTable1")
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act
    Dim accFinder As AccountsFinder: Set accFinder = New AccountsFinder
    Call accFinder.Initialize(cb)
    Dim dic As Dictionary
    Set dic = accFinder.FindAccounts(ofReportingUnit:="���k�u���b�N�u�K��")
    'Assert
    Call KzCls
    Dim key As Variant
    For Each key In dic
        Debug.Print key & ":" & dic(key)
    Next
    
'�x�o/������/�ʐM��:17
'�x�o/���Ɣ�/�w�p��:9
'�x�o/���Ɣ�/�ʐM��:1

    'TearDown
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
    wb.Close
End Sub





