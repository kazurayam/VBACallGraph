Attribute VB_Name = "BbUtil"
Option Explicit

'BbUtil


Public Function VarTypeAsString(ByVal var As Variant) As String
    ' 引数varのtypeを調べて変数の型を示す文字列（"Integer"など）を返す
    Dim typeValue As Long: typeValue = VarType(var)
    Dim result As String: result = "unknown"
    If typeValue = 2 Then
        result = "Integer"
    ElseIf typeValue = 3 Then
        result = "Long"
    ElseIf typeValue = 5 Then
        result = "Double"
    ElseIf typeValue = 8 Then
        result = "String"
    ElseIf typeValue = 11 Then
        result = "Boolean"
    ElseIf typeValue = 7 Then
        result = "Date"
    ElseIf typeValue = 9 Then
        result = "Object"
    ElseIf typeValue = 0 Then
        result = "Variant"
    ElseIf typeValue = 8200 Then
        result = "String()"
    ElseIf typeValue = 8194 Then
        result = "Integer()"
    Else
        result = Str(typeValue)
    End If
    VarTypeAsString = result
End Function


Public Function ResolveExternalFilePath( _
        ByVal theWorkbook As Workbook, _
        ByVal sheetName As String, _
        ByVal rangeLiteral As String) As String
    'theWorkbookとして与えられたワークブックのなかに
    'sheetNameとして与えられたワークシートがあって、その中に
    'rangeLiteralとして与えられたセルがあって、そのなかに
    '外部ファイルのパスが書いてあると期待する。
    'そのパスがtheWorkbookを基底とする相対パスであると期待する。
    '外部ファイルのパスを発見し、それを絶対パスに変換して、Functionの値として返す。
    'この関数は.xlsmファイルの可搬性を高めるのに有用である。
    '.xlsmファイルから見た外部ファイルのパスをVBAコードのなかに
    '固定値として書くのではなく、
    'ワークシートのセルの値として書くことを可能にする。
    If BbWorksheet.IsWorksheetPresentInWorkbook(theWorkbook, sheetName) Then
        Dim ws As Worksheet: Set ws = theWorkbook.Worksheets(sheetName)
        Dim path As String
        path = ws.Range(rangeLiteral)
        ResolveExternalFilePath = BbFile.AbsolutifyPath(BbFile.ToLocalFilePath(theWorkbook.path), path)
    Else
        Debug.Print theWorkbook.FullName + " does not have a worksheet named " + sheetName
        ResolveExternalFilePath = ""
    End If
End Function


