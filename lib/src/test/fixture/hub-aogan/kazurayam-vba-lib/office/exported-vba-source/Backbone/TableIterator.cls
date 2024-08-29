VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' TableIteratorクラス

' じつはVBAにはListObjectという組み込みオブジェクトがある。
' ところが浦山和昭はVBAの素人なので、はじめListObjectを知らなかった。
' 必要に迫られてListObjectのようなものを自作してしまった。
' それがTableIteratorとISheetAccessorとSheetAccessor

' Worksheetの中にTableがあるとイメージする。
' Tableとは一番上の行にHeader行があって、Header下にデータの行があるとというイメージ。
' データの行をKeyとValueのペアつまりDictionariオブジェクトとして順番に取り出すメソッドを提供する。
' そのメソッドの実装をここに隠すしてCall元のコードの見通しをよくすることを目的とする。
' |||||
' ||name   |address   |age|
' ||Alice  |here      | 16|
' ||Bob    |there     | 23|
' ||Charlie|everywhere|   |
' ||Dave   |          |  6|
' |||||

Dim shAccess As ISheetAccessor
Dim sheet As Worksheet
Dim headerRowIndex As Long  ' 2 in the above example
Dim headerColLeft As Long   ' 2
Dim headerColRight As Long  ' 4
Dim keys() As String        '
Dim dataRowTop As Long      ' 3
Dim dataRowBottom As Long   ' 6

Dim currentRowIndex As Long
Dim arrUtil As ArrayUtil
    
Public Sub Class_Initialize()
    Set shAccess = New SheetAccessor
    Set arrUtil = New ArrayUtil
End Sub

' Worksheetオブジェクトと、見出し行の行番号と、見出し行の開始列番号をCall元が与える。
' 見出し行に含まれるKeyの集合をInitializeが探しだす。
' 見出し行が1行だけであることを暗黙的に前提する。
' 見出し行の次からデータ行が始まると前提する。
' データ行が全部でいくつあるかはInitializeが探し出す。
Public Sub Initialize( _
        ByVal sheet_ As Worksheet, _
        ByVal headerRowIndex_ As Long, _
        ByVal headerColLeft_ As Long)
    Set sheet = sheet_
    headerRowIndex = headerRowIndex_
    headerColLeft = headerColLeft_
    headerColRight = shAccess.FindLastColumn(sheet, headerRowIndex, headerColLeft)
    
    'Debug.Print "sheet.Name=" & sheet.Name
    'Debug.Print "headerRowIndex=" & headerRowIndex
    'Debug.Print "headerColLeft=" & headerColLeft
    'Debug.Print "headerColRight=" & headerColRight
    
    keys = shAccess.ReadSingleRow(sheet, headerRowIndex, headerColLeft, headerColRight)
    
    ' for デバッグ
    Call arrUtil.PrintArray1D(keys)
    
    dataRowTop = headerRowIndex + 1
    dataRowBottom = shAccess.GetRowBottom(sheet, headerColLeft)
    currentRowIndex = dataRowTop
End Sub

' CallerがTableIteratorインスタンスのNextDictionary()メソッドを
' 1回CallするたびにNextDictionary()メソッドは次のデータ行1つを
' Dictionaryオブジェクトに変換して返す。
' CallerはNextDictionary()を呼ぶ前にHasNext()メソッドを呼んで
' Trueが返ってくることを確かめるべきである。
' すべてのデータ行を取り扱ったあとでさらにCallされたらNextDictionary()
' は警告メッセージを表示し、結果として空のDictionaryを返す。
Public Function NextDictionary() As Dictionary
    Dim i As Integer
    If HasNext() = True Then
        
        Dim dict As Dictionary: Set dict = New Dictionary
        For i = LBound(keys) To UBound(keys)
            
            'Debug.Print "currentRowIndex=" & currentRowIndex & _
            '            ", i=" & i & _
            '            ", key=" & keys(i) & _
            '            ", value=" & shAccess.GetCellValue(sheet, currentRowIndex, 1 + i)
            
            If keys(i) <> "" Then
                If Not dict.Exists(keys(i)) Then
                    dict.Add keys(i), shAccess.GetCellValue(sheet, currentRowIndex, 2 + i) ' なんで1+iかというと、keys配列の添字が1で始まるから
                Else
                    ' 同じkeyが2回以上現れたら先のほうを採用し後のほうを無視する
                End If
            End If
        Next i
        
        ' increment the index
        currentRowIndex = currentRowIndex + 1
        
        Set NextDictionary = dict
    Else
        Debug.Print "HasNext() returned False"
        Set NextDictionary = New Dictionary
    End If
End Function

' つぎにNextDictionary()メソッドが呼ばれたときに有意なデータを返せる状態ならばHasNext()はTrueを返す。
' さもなければ（つまりもう最終のデータまで処理済みならば）HasNext()はFalseを返す。
Public Function HasNext() As Boolean
    If currentRowIndex <= dataRowBottom Then
        HasNext = True
    Else
        HasNext = False
    End If
End Function




