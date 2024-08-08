VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISheetAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' SheetAccessorクラスの外部仕様を規定するインタフェースです

' 指定されたセルを左端とする見出し領域の最終列の番号を返す
Public Function FindLastColumn( _
    ByVal sheet As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colLeft As Long) As Long
End Function

' 指定された列の最終行を取得
Public Function GetRowBottom( _
    ByVal sheet As Worksheet, _
    ByVal colLeft As Long) As Long
End Function

' 指定された範囲の値を2次元配列で取得
Public Function ReadMatrix( _
    ByVal sheet As Worksheet, _
    ByVal rowTop As Long, _
    ByVal colLeft As Long, _
    ByVal rowBottom As Long, _
    ByVal colRight As Long) As Variant
End Function

' １つの行のセル(Row,Col1)とセル(Row,Col2)によって指定された範囲のセル値を1次元配列として取得
Public Function ReadSingleRow( _
    ByVal sheet As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colLeft As Long, _
    ByVal colRight As Long) As Variant
End Function

' 指定された範囲に2次元配列の値を書き込む
Public Sub WriteMatrix( _
    ByVal sheet As Worksheet, _
    ByVal rowTop As Long, _
    ByVal colLeft As Long, _
    ByVal rowBottom As Long, _
    ByVal colRight As Long, _
    ByVal matrix As Variant)
End Sub

' (行i, 列i)のセルのテキストを返す
Public Function GetCellValue( _
    ByVal sheet As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colIndex As Long) As String
End Function

' (行i, 列j)にセルが存在していたらTrue、無ければFalseを返す
Public Function CellExists( _
    ByVal sheet As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colIndex As Long) As Boolean
End Function





