Attribute VB_Name = "Helpers"
Option Explicit

' Cashbook�ɋL�ڂ��ꂽ���o���f�[�^������Ȗڂ��Ƃɕ��ނ������𐔂��ĕ\������
Public Sub PrintAccounts(ByVal cb As Cashbook, ByVal rpUnit As String, ByVal positiveLike As Boolean)
    'Debug.Print "PrintAccounts was called"
    Dim accFinder As AccountsFinder: Set accFinder = CreateAccountsFinder(cb)
    Call accFinder.Initialize(cb, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    'Debug.Print "initialized AccountFinder"
    Dim dic As Dictionary: Set dic = accFinder.FindAccounts(rpUnit, positiveLike)
    Debug.Print "dic was created"
    Debug.Print "dic.Count=" & dic.Count
    ' dic�̃L�[��String�A���Ƃ��΁u�x�o/������/�ʐM��v
    ' dic��value��CashList�I�u�W�F�N�g
    
    ' �L�[�̈ꗗ��z��ɂƂ肾���ă\�[�g����
    Dim var As Variant
    var = dic.Keys
    Call InsertionSort(var, LBound(var), UBound(var))
    
    ' �L�[���Ƃɖ��ׂ�Print����
    Dim cList As CashList
    Dim msg As String
    Dim key As Variant
    For Each key In var
        Set cList = dic(key)
        Debug.Print key,
        Debug.Print cList.Count & "��",
        If key Like "����/*" Then
            Debug.Print cList.SumOfIncomeAmount() & " �~"
        ElseIf key Like "�x�o/*" Then
            Debug.Print cList.SumOfExpenseAmount() & " �~"
        Else
            Debug.Print ""
        End If
    Next
End Sub
