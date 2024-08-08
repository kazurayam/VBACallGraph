Attribute VB_Name = "KVLプロシージャー一覧を作る"
Option Explicit

Private Sub KVLプロシージャー一覧を作る()

    Dim dicProcInfo As New Dictionary
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim i As Long
  
    'ブックの全モジュールを処理
    With wb.VBProject
        For i = 1 To .VBComponents.Count
            Call getCodeModule.getCodeModule(dicProcInfo, wb, .VBComponents(i).Name)
        Next
    End With
  
    'Dictionaryよりシートに出力
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("プロシージャー一覧")
    Dim v
    With ws
        .Cells.Clear
        .Range("A1:G1").Value = Array("プロシージャー in " & ThisWorkbook.Name, "モジュール", "スコープ", "種別", "行位置", "ソース", "コメント")
        i = 2
        For Each v In dicProcInfo.Items
            .Cells(i, 1) = v.ProcName
            .Cells(i, 2) = v.ModName
            .Cells(i, 3) = v.Scope
            .Cells(i, 4) = v.ProcKindName
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
    
    Set dicProcInfo = Nothing
End Sub
