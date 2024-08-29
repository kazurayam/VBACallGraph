Attribute VB_Name = "Test_DictionaryUtil"
Option Explicit

Public Sub Test_printDictionary()
    Call KzUtil.KzCls
    Dim dict As Dictionary: Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "first", 30
    dict.Add "second", 40
    dict.Add "third", 100
    Dim dictUtil As DictionaryUtil: Set dictUtil = New DictionaryUtil
    Call dictUtil.PrintDictionary(dict)
End Sub


