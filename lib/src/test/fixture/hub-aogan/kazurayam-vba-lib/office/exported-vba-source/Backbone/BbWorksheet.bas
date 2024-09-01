Attribute VB_Name = "BbWorksheet"
Option Explicit

'KzWorksheet


' 指定されたワークブックのなかに指定された名前のシートが存在していたらTrueを返す
Public Function IsWorksheetPresentInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            flg = True
            Exit For
        End If
    Next ws
    IsWorksheetPresentInWorkbook = flg
End Function


' 指定されたワークブックのなかに指定された名のシートが存在しなければ追加する
' 追加したときはTrueを返す。
' シートがすでにあったならばなにもせずFalseを返す
Public Function CreateWorksheetInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim flg As Boolean: flg = False
    If Not IsWorksheetPresentInWorkbook(wb, sheetName) Then
        Dim ws As Worksheet
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = sheetName
        flg = True
    End If
    CreateWorksheetInWorkbook = flg
End Function


' 指定されたワークブックのなかに指定された名のシートが存在すれば削除する
' 削除したときはTrueを返す。
' シートが無ければなにもせずFalseを返す
Public Function DeleteWorksheetInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    '指定されたブックに指定したシートが存在するかチェック
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            'あればシートを削除する
            Application.DisplayAlerts = False    ' メッセージを非表示
            ws.Delete
            Application.DisplayAlerts = True
            flg = True
            Exit For
        End If
    Next ws
    DeleteWorksheetInWorkbook = flg
End Function


' コピー元として指定されたワークブックのワークシートを
' コピー先として指定されたワークブックのワークシートにコピーする。
' @param sourceWorkbook コピー元のWorkbook
' @param sourceSheetName コピー元のWorksheetの名前
' @param targetWorkbook コピー先のWorkbook
' @param targetSheetName コピー先のWorksheetの名前
'
' sourceWorkbookとsourceSheetNameで示されるワークシートが存在していることが必要。さもなければエラーになる。
'
' targetWorkbookのなかにtargetSheetNameで示されるワークシートが未だ無い場合とすでに在る場合とがありうる。
' まだ無ければsourceのシートをコピーすることでtargetSheetNameのワークシートが新しくできる。
' すでに在ったらtargetWorkbookのなかの古いシートを削除して、sourceのシートをコピーする。
' ただしsourceWorkbookとtargetWorkbookが同じで、かつ、sourceSheetNameとtargetSheetNameが同じ場合は
' 指定ミスだからエラーとする。
'
Public Sub FetchWorksheetFromWorkbook(ByVal sourceWorkbook As Workbook, _
                                        ByVal sourceSheetName As String, _
                                        ByVal targetWorkbook As Workbook, _
                                        ByVal targetSheetName As String)
'エラーが起きたときはErrorHandlerに跳ぶ
On Error GoTo ErrorHandler

    'sourceとtargetが同じ場合はエラーとする
    If sourceWorkbook.path = targetWorkbook.path And sourceSheetName = targetSheetName Then
        Err.Raise Number:=2022, Description:="同じワークブックの同じワークシートをsourceとtargetに指定してはいけません"
    End If
    
    '標的のワークシートがすでにターゲットのワークブックにあったら削除する
    If IsWorksheetPresentInWorkbook(targetWorkbook, targetSheetName) Then
        Application.DisplayAlerts = False
        targetWorkbook.Worksheets(targetSheetName).Delete
        Application.DisplayAlerts = True
    End If
    
    'コピー元ワークシートのすべてのセルをコピーして
    '新しいワークシートとしてターゲットのワークブックに挿入する
    sourceWorkbook.Worksheets(sourceSheetName).Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
    
    '新しいワークシートの名前を指定されたように変更する
    ActiveSheet.Name = targetSheetName
        
    
 '例外処理
ErrorHandler:
    ' もしもエラーが起きていたならcall元に伝播させる
    If Err.Number <> 0 Then
        Call BbLog.Info("BbWorksheet", "FetchWorksheetFromWorkbook", "sourceWorkbook : " & sourceWorkbook.FullName)
        Call BbLog.Info("BbWorksheet", "FetchWorksheetFromWorkbook", "sourceSheetName: " & sourceSheetName)
        Call BbLog.Info("BbWorksheet", "FetchWorksheetFromWorkbook", "targetWorkbook : " & targetWorkbook.FullName)
        Call BbLog.Info("BbWorksheet", "FetchWorksheetFromWorkbook", "targetSheetName: " & targetSheetName)
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If

End Sub

