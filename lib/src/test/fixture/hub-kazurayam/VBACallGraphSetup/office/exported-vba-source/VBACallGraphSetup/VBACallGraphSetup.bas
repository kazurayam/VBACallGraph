Attribute VB_Name = "VBACallGraphSetup"
Option Explicit


Public Sub ExportModules(ByVal wb As Workbook)

    Dim dicProcInfo As New Dictionary
    Dim i As Long
  
    'ブックの全モジュールを処理
    With wb.VBProject
        For i = 1 To .VBComponents.Count
            Call CodeModuleUtil.GetCodeModule(dicProcInfo, wb, .VBComponents(i).Name, .VBComponents(i).Type)
        Next
    End With
  
    '出力先としてのワークシートを準備する
    Dim sheetName As String: sheetName = "ExportedModules"
    Dim r As Boolean
    r = KzCreateWorksheetInWorkbook(wb, sheetName)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    
    'プロシジャの情報をシートに出力する
    Dim v
    With ws
        .Cells.Clear
        .Range("A1:I1").Value = Array("Project", "ModuleType", "Module", "Scope", "ProcKind", "Procedure", "LineNo", "Source", "Comment")
        .Range("A1:H1").Interior.Color = RGB(200, 200, 200) ' 背景色をグレー
        i = 2
        For Each v In dicProcInfo.Items
            .Cells(i, 1) = wb.VBProject.Name   ' KazurayamVbaLib
            .Cells(i, 2) = v.ModuleType        ' Standard | Class (Sheet Module, ThisWorkbook Module, Userformsは未サポート)
            .Cells(i, 3) = v.ModName           ' KzSensible
            .Cells(i, 4) = v.Scope             ' Public | Private | Static
            ' .Cells(i, 5) = v.ProcKindName
            .Cells(i, 5) = FormatProcKindName(v.ProcKindName, v.Source)
            .Cells(i, 6) = v.ProcName          ' KzProcedureList
            .Cells(i, 7) = v.LineNo
            .Cells(i, 8) = v.Source
            .Cells(i, 9) = "'" & v.Comment
            i = i + 1
        Next
        Cells.EntireRow.AutoFit
        Cells.EntireColumn.AutoFit
        Range("F1").ColumnWidth = 30
        Range("H1:I1").ColumnWidth = 100
        
    End With

    'シートの行をプロジェクト名>モジュール名>プロシジャ名の昇順でソートする
    With ws.Sort
        With .SortFields
            .Clear
            .Add Key:=ws.Range("A2"), Order:=xlAscending
            .Add Key:=ws.Range("C2"), Order:=xlAscending
            .Add Key:=ws.Range("F2"), Order:=xlAscending
        End With
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(i, 9))
        .Header = xlYes
        .Apply
    End With
    
    '行の高さを自動調節する
    ws.Rows.AutoFit
    
    Set dicProcInfo = Nothing
End Sub

' 指定されたワークブックのなかに指定された名のシートが存在しなければ追加する
' 追加したときはTrueを返す。
' シートがすでにあったならばなにもせずFalseを返す
Private Function KzCreateWorksheetInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim flg As Boolean: flg = False
    If Not KzIsWorksheetPresentInWorkbook(wb, sheetName) Then
        Dim ws As Worksheet
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = sheetName
        flg = True
    End If
    KzCreateWorksheetInWorkbook = flg
End Function

' 指定されたワークブックのなかに指定された名前のシートが存在していたらTrueを返す
Public Function KzIsWorksheetPresentInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            flg = True
            Exit For
        End If
    Next ws
    KzIsWorksheetPresentInWorkbook = flg
End Function


Private Function FormatProcKindName(ByVal ProcKindName As String, ByVal Source As String) As String
    If (InStr(1, LCase(Source), "function ") > 0) Then
        FormatProcKindName = "Function"
    ElseIf (InStr(1, LCase(Source), "sub ") > 0) Then
        FormatProcKindName = "Sub"
    Else
        FormatProcKindName = ProcKindName  ' Property Let | Property Set | Property Get
    End If
End Function
