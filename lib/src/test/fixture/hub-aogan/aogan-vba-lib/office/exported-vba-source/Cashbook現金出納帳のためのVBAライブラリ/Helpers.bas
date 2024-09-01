Attribute VB_Name = "Helpers"
Option Explicit

' Helpers

' Cashbook�ɋL�ڂ��ꂽ���o���f�[�^������Ȗڂ��Ƃɕ��ނ������𐔂��ĕ\������
Public Sub PrintAccounts(ByVal cb As Cashbook, ByVal rpUnit As String, ByVal positiveLike As Boolean)
    'Debug.Print "PrintAccounts was called"
    Dim cbTransformer As CashbookTransformer
    Set cbTransformer = Factories.CreateCashbookTransformer(cb)
    Call cbTransformer.Initialize(cb, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    'Debug.Print "initialized AccountFinder"
    Dim cashListDic As Dictionary
    Set cashListDic = cbTransformer.ByAccounts(rpUnit, positiveLike)
    Debug.Print "cashListDic.Count=" & cashListDic.Count
    ' dic�̃L�[��String�A���Ƃ��΁u�x�o/������/�ʐM��v
    ' dic��value��CashList�I�u�W�F�N�g
    
    ' �L�[�̈ꗗ��z��ɂƂ肾���ă\�[�g����
    Dim accounts As Variant
    accounts = cashListDic.Keys
    Call BbArraySort.InsertionSort(accounts, LBound(accounts), UBound(accounts))
    
    ' �L�[���Ƃɖ��ׂ�Print����
    Dim account_ As Variant
    Dim cashList_ As CashList
    Dim msg As String
    For Each account_ In accounts
        Set cashList_ = cashListDic(account_)
        Debug.Print account_,
        Debug.Print cashList_.Count & "��",
        If account_ Like "����/*" Then
            Debug.Print cashList_.SumOfIncomeAmount() & " �~"
        ElseIf account_ Like "�x�o/*" Then
            Debug.Print cashList_.SumOfExpenseAmount() & " �~"
        Else
            Debug.Print ""
        End If
    Next
End Sub
