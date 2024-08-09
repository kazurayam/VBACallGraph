Attribute VB_Name = "Test_TableIterator"
Option Explicit


' TableIterator�N���X�����j�b�g�e�X�g����

Public Sub Test_Initialize()
    Call G.Cls
    Dim sheet As Worksheet: Set sheet = Worksheets("Sheet1")
    Dim tblIter As TableIterator: Set tblIter = New TableIterator
    Call tblIter.Initialize(sheet, 2, 2)
    
    ' ���ʂ̃A�T�[�V����
    Debug.Assert tblIter.HasNext = True
End Sub

Public Sub Test_HasNext_NextDictionary()
    Call G.Cls
    Dim sheet As Worksheet: Set sheet = Worksheets("Sheet1")
    Dim tblIter As TableIterator: Set tblIter = New TableIterator
    Call tblIter.Initialize(sheet, 2, 2)

    Dim dict As Dictionary
    Dim dictUtil As DictionaryUtil: Set dictUtil = New DictionaryUtil
    
    Do While tblIter.HasNext()
        Set dict = tblIter.nextDictionary
        Debug.Print "-------------------------------"
        Call dictUtil.printDictionary(dict)
        Debug.Assert Len(dict.item("���O")) > 0
    Loop
End Sub


