Attribute VB_Name = "KzMetaprogramming"
Option Explicit


Public Sub KzProcedureList(ByVal wb As Workbook)

    Dim dicProcInfo As New Dictionary
    Dim i As Long
  
    'ブックの全モジュールを処理
    With wb.VBProject
        For i = 1 To .VBComponents.Count
            Call getCodeModule.getCodeModule(dicProcInfo, wb, .VBComponents(i).Name)
        Next
    End With
  
    '出力先としてのワークシートを準備する
    Dim sheetName As String: sheetName = "プロシージャ一覧"
    Dim r As Boolean
    r = KzCreateWorksheetInWorkbook(wb, sheetName)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    
    'プロシージャーの情報をシートに出力する
    Dim v
    With ws
        .Cells.Clear
        .Range("A1:G1").Value = Array(wb.Name, "モジュール", "スコープ", "種別", "行位置", "ソース", "コメント")
        .Range("A1:G1").Interior.Color = RGB(200, 200, 200) ' 背景色をグレーに
        i = 2
        For Each v In dicProcInfo.Items
            .Cells(i, 1) = v.ProcName
            .Cells(i, 2) = v.ModName
            .Cells(i, 3) = v.Scope
            ' .Cells(i, 4) = v.ProcKindName
            .Cells(i, 4) = IIf(InStr(1, LCase(v.Source), " function ") > 0, "Function", "Sub")
            .Cells(i, 5) = v.LineNo
            .Cells(i, 6) = v.Source
            .Cells(i, 7) = "'" & v.Comment
            i = i + 1
        Next
        Cells.EntireRow.AutoFit
        Cells.EntireColumn.AutoFit
    End With

    'シートの行をプロシージャ名の昇順でソートする
    With ws.Sort
        With .SortFields
            .Clear
            .Add key:=ws.Range("A1"), Order:=xlAscending
        End With
        .SetRange ws.Range(Cells(1, 1), Cells(i, 7))
        .Header = xlYes
        .Apply
    End With
    
    '行の高さを自動調節する
    ws.Rows.AutoFit
    
    Set dicProcInfo = Nothing
End Sub

