Attribute VB_Name = "Test_SheetAccessor"
Option Explicit


' SheetAccessorクラスをユニットテストする

Public Sub Test_FindLastColumn()
    Dim sheet As Worksheet: Set sheet = Worksheets("Sheet1")
    Dim accessor As ISheetAccessor
    Set accessor = New SheetAccessor
    Dim col As Long: col = accessor.FindLastColumn(sheet, 2, 2)
    
    ' 結果のアサーション
    Call G.Cls
    'Debug.Print "col=" & col
    Debug.Assert col = 4
End Sub

Public Sub Test_ReadMatrix()
    Call G.Cls
    Dim sheet As Worksheet: Set sheet = Worksheets("Sheet1")
    Dim accessor As ISheetAccessor: Set accessor = New SheetAccessor
    Dim arrUtil As ArrayUtil: Set arrUtil = New ArrayUtil
    
    Dim arr2D As Variant
    arr2D = accessor.ReadMatrix(sheet, 2, 2, 5, 4)
    Debug.Print "arr2D is:"
    Call arrUtil.PrintArray2D(arr2D)
End Sub

Public Sub Test_ReadSingleRow()
    Call G.Cls
    Dim sheet As Worksheet: Set sheet = Worksheets("Sheet1")
    Dim accessor As ISheetAccessor: Set accessor = New SheetAccessor
    Dim arrUtil As ArrayUtil: Set arrUtil = New ArrayUtil
    Dim keys As Variant
    keys = accessor.ReadSingleRow(sheet, rowIndex:=2, colLeft:=2, colRight:=4)
    Debug.Print "keys array:"
    Call arrUtil.PrintArray1D(keys)
End Sub

Public Sub Test_GetCellValue()
    Call G.Cls
    Dim sheet As Worksheet: Set sheet = Worksheets("Sheet1")
    Dim accessor As ISheetAccessor: Set accessor = New SheetAccessor
    Debug.Assert accessor.GetCellValue(sheet, 3, 4) = "A"
End Sub

Public Sub Test_CellExists_True()
    Call G.Cls
    Dim sheet As Worksheet: Set sheet = Worksheets("Sheet1")
    Dim accessor As ISheetAccessor: Set accessor = New SheetAccessor
    Debug.Assert accessor.CellExists(sheet, 2, 2)   ' 「名前」のセル
End Sub

Public Sub Test_CellExists_False()
    Call G.Cls
    Dim sheet As Worksheet: Set sheet = Worksheets("Sheet1")
    Dim accessor As ISheetAccessor: Set accessor = New SheetAccessor
    Debug.Assert Not accessor.CellExists(sheet, 5, 3)   ' 徳川家康の勤務先はシートに書いてない
End Sub



