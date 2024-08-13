VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' Worksheetを操作するメソッドをまとめたクラスです
'
' 行番号、列番号は１から始まる

Implements ISheetAccessor

' 指定された行番号と列番号が指すセルを左端とするRangeつまり空でないセルの連鎖領域の右端のセルの番号を返す。
' この関数はシートのなかの「見出し行」が持っている列数を知るために用いる
Public Function ISheetAccessor_FindLastColumn( _
    ByVal sheet As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colLeft As Long) As Long
    ISheetAccessor_FindLastColumn = _
        sheet.Cells(rowIndex, colLeft).End(xlToRight).Column
End Function

' 第１列から指定された列番号までの範囲にデータがある最大の行番号を取得
' この関数はシートのなかで有効なデータの行が最大いくつかを知るために用いる
Public Function ISheetAccessor_GetRowBottom( _
    ByVal sheet As Worksheet, _
    ByVal col As Long) As Long
    ISheetAccessor_GetRowBottom = _
        sheet.Cells(sheet.Rows.Count, col).End(xlUp).Row
End Function

' シートのセル(Row1,Col1)とセル(Row2, Col2)によって指定されたRangeオブジェクトの
' Valueを返す。返されるオブジェクトは2次元配列である。
' RangeオブジェクトのValueは特徴的である。すなわち
' Valueの第1次元のLBoundは必ず1でUBoundは必ず2である。
' Valueの第2次元のLBoundは必ず1でUBoundは1以上の整数である。
'
Public Function ISheetAccessor_ReadMatrix( _
    ByVal sheet As Worksheet, _
    ByVal rowTop As Long, _
    ByVal colLeft As Long, _
    ByVal rowBottom As Long, _
    ByVal colRight As Long) As Variant
    ISheetAccessor_ReadMatrix = _
        sheet.Range(sheet.Cells(rowTop, colLeft), _
                    sheet.Cells(rowBottom, colRight)).Value
End Function

' シートのセル(Row,Col1)とセル(Row,Col2)によって指定されたRangeオブジェクトのValueのなかから
' 1行目のarrayを取り出して返す。
' 返される1次元配列のLBoundはCol1に等しく、UBoundはCol2に等しい。
Public Function ISheetAccessor_ReadSingleRow( _
    ByVal sheet As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colLeft As Long, _
    ByVal colRight As Long) As Variant
    
    Dim arrUtil As ArrayUtil: Set arrUtil = New ArrayUtil
    
    ' array2Dは2次元配列
    Dim array2D As Variant
    array2D = ISheetAccessor_ReadMatrix(sheet, rowIndex, colLeft, rowIndex, colRight)
    
    ' array1Dは1次元配列
    Dim array1D() As String
    ReDim array1D(colRight)
    
    ' array2Dの1行目を取り出して1次元配列を得る
    Dim i As Long
    For i = 1 To colRight - colLeft + 1
        array1D(i) = array2D(1, i)   ' 1行目を取り出す
    Next
    
    ' 戻り値をセット
    ISheetAccessor_ReadSingleRow = array1D
    
End Function

' 指定されたセル範囲に2次元配列の値を書き込む
Public Sub ISheetAccessor_WriteMatrix( _
    ByVal sheet As Worksheet, _
    ByVal rowTop As Long, _
    ByVal colLeft As Long, _
    ByVal rowBottom As Long, _
    ByVal colRight As Long, _
    ByVal matrix As Variant)
    sheet.Range(sheet.Cells(rowTop, colLeft), _
                sheet.Cells(rowBottom, colRight)).Value = matrix
End Sub

' 指定された行番号と列番号に該当するセルに表示されたテキストを返す
Public Function ISheetAccessor_GetCellValue( _
    ByVal sheet As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colIndex As Long) As String
    Dim strBuf As String
    If ISheetAccessor_CellExists(sheet, rowIndex, colIndex) = True Then
        strBuf = sheet.Cells(rowIndex, colIndex).Value
        strBuf = Trim(strBuf)                   ' 前後の空白を削除
        strBuf = Replace(strBuf, vbLf, "")      ' 改行を削除
        ' strBuf = StrConv(strBuf, vbNarrow)    ' 全角を半角に
        ' strBuf = StrConv(strBuf, vbUpperCase) ' 英小文字を大文字に
    Else
        strBuf = "not exist"
    End If
    ISheetAccessor_GetCellValue = strBuf
End Function

' (i,j)に該当するセルが存在するかどうかを調べてTrue/Falseを返す
' でも、セルの有無を調べることは無理で、取り出したTextが空文字列""かどうかで判断する
Public Function ISheetAccessor_CellExists( _
        ByVal sheet As Worksheet, _
        ByVal rowIndex As Long, _
        ByVal colIndex As Long) As Boolean
    Dim myRange As Range
    Dim t As String
    
    On Error Resume Next
    Set myRange = sheet.Cells(rowIndex, colIndex)
    ' もし(rowIndex,colIndex)にセルが存在していなかったらここでErrorが発生する
    ' しかしOn Error Resume Nextと書いてあるので、次の行に進む。
    t = myRange.Value
    ' tに空じゃない文字があればなにか在ったとみなす。
    ' tが空文字ならばなにも無かったとみなす
    On Error GoTo 0
    ' Resume Nextを解除しておく
    If t <> "" Then
        ISheetAccessor_CellExists = True
    Else
        ISheetAccessor_CellExists = False
    End If
End Function





