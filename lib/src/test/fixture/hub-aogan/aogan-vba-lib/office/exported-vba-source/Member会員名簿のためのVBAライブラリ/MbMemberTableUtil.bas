Attribute VB_Name = "MbMemberTableUtil"
Option Explicit

'
' MbMemberTableUtil --- 会員名簿のExcelワークシートを扱う共用のSubとFunctionを提供する
'

' 会員名簿のExcelファイルのパスをパラメータとして受ける。
' そのなかの会員名簿ワークシートの名前をmemberSheetNameパラメータとして受ける。
' ワークシートを読み取って、指定されたワークブックのなかにコピーする。
' 出力先としてのワークシートの名前を指定しなければデフォルトとして「work会員名簿」とする。
' そして出力されたワークシートにに含まれているExcel Tableの内容ListObjectに変換して返す。
Public Function FetchMemberTable(memberFilePath As String, _
        memberSheetName As String, _
        targetWorkbook As Workbook, _
        Optional ByVal targetSheetName As String = "work会員名簿", _
        Optional ByVal renew As Boolean = True, _
        Optional ByVal tableId As String = "MembersTable13") As ListObject


    ' targetワークブックのなかに　"work会員名簿" シートが無ければ作る。
    ' すでにコピーが存在していてかつrenewがTrueと指定されていたらコピーを上書きする。
    ' すでにコピーが存在していてかつrenewがFalseならば何もしない。
    If (Not BbWorksheet.IsWorksheetPresentInWorkbook(targetWorkbook, targetSheetName)) Or _
        (BbWorksheet.IsWorksheetPresentInWorkbook(targetWorkbook, targetSheetName) And renew = True) Then
        
        ' 外部にある会員名簿ワークブックのウインドウを開く
        Dim sourceWorkbook As Workbook
        ' memberFilePathがtargetWorkbookのPathを基底とする相対パスで指定されていても大丈夫なように用心する
        Set sourceWorkbook = Workbooks.Open(BbFile.AbsolutifyPath( _
                                            targetWorkbook.Path, memberFilePath))
                                            
        '別のワークブックをopenするとThisWorkbookが自動的に切り替わってしまう事に注意せよ
        
        ' 外部にある会員名簿ExcelファイルのワークシートをカレントのWookbookにコピーする
        Call BbWorksheet.FetchWorksheetFromWorkbook( _
                sourceWorkbook, memberSheetName, _
                targetWorkbook, targetSheetName)
        
        ' 外部にある会員名簿ワークブックのwindowを閉じる。閉じないで放っておいてはいけません。
        ' 人が手動でウインドウを閉じなければならなくず、厄介だ。発つ鳥 跡を濁さず。
        Application.DisplayAlerts = False   '「変更内容を保存しますか」ダイアログを表示しない
        sourceWorkbook.Close
        Application.DisplayAlerts = True
    End If
    
    ' targetワークブックの”work会員名簿”シートのなかにExcel Tableがあるはず
    Dim ws As Worksheet: Set ws = targetWorkbook.Worksheets(targetSheetName)
    
    ' テーブルをListObjectに変換する
    Dim tbl As ListObject: Set tbl = ws.ListObjects(tableId)
    
    ' ListObjectの形を整える --------------------------------------------------
    
    '列幅を調整
    tbl.ListColumns("氏名カナ").Range.EntireColumn.AutoFit
    'tbl.ListColumns("勤務先名").Range.EntireColumn.AutoFit
    'tbl.ListColumns("異動").Range.EntireColumn.AutoFit
    
    '重要でない列を非表示にする。シートを印刷したときに便利なように
    'tbl.ListColumns("年齢").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("医登No").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("医登録日").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("日眼医登録").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("勤務〒").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("勤務先住所").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("勤先TELNo").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("自宅〒").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("自宅住所").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("自宅TELNo").Range.EntireColumn.Hidden = True
    'tbl.ListColumns("携帯番号").Range.EntireColumn.Hidden = True
    
    Set FetchMemberTable = tbl
        
End Function


