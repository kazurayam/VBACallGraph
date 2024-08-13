Attribute VB_Name = "LearningModule"
Option Explicit

' VBAをコーディングするうえで疑問に思ったことを解決するため
' いろいろ試した。そのコードをここに配置する。

Sub ListWorksheetsInAoganCashbook()
    ' 青森県眼科医会の現金出納帳のワークブックを開き、
    ' そのなかにあるワークシート名をDebug.Printする
    Dim wb As Workbook
    Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Dim i As Long
    Call KzCls
    For i = 1 To wb.Worksheets.Count
        Debug.Print wb.Worksheets(i).Name
    Next i
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
End Sub



