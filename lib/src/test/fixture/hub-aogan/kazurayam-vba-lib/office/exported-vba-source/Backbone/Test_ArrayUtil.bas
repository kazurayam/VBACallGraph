Attribute VB_Name = "Test_ArrayUtil"
Option Explicit

Public Sub Test_PrintArray2D()
    Call KzUtil.KzCls
    Dim sheet As Worksheet: Set sheet = Worksheets("Sheet1")
    Dim accessor As ISheetAccessor: Set accessor = New SheetAccessor
    Dim arrUtil As ArrayUtil: Set arrUtil = New ArrayUtil
    Dim arr2D As Variant
    arr2D = accessor.ReadMatrix(sheet, 3, 2, 5, 4)
    Debug.Print "arr2D is:"
    Call arrUtil.PrintArray2D(arr2D, 3, 2)
    Debug.Print
    Call arrUtil.PrintArray2D(arr2D)
End Sub

Public Sub Test_PrintArray1D()
    Call KzUtil.KzCls
    Dim arr1D As Variant
    arr1D = Array("foo", "bar", "baz")
    Dim arrUtil As ArrayUtil: Set arrUtil = New ArrayUtilz
    Debug.Print "arr1D is:"
    Call arrUtil.PrintArray1D(arr1D)
End Sub

Public Sub Test_PrintArray1D_startsWith1()
    Call KzUtil.KzCls
    Dim arr1D(3) As Variant
    arr1D(1) = "foo"
    arr1D(2) = "bar"
    arr1D(3) = "baz"
    Dim arrUtil As ArrayUtil: Set arrUtil = New ArrayUtil
    Debug.Print "arr1D is:"
    Call arrUtil.PrintArray1D(arr1D)
End Sub


