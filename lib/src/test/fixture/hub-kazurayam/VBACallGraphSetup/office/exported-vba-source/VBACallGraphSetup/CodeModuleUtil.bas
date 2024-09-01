Attribute VB_Name = "CodeModuleUtil"
Option Explicit


'https://www.excel-ubara.com/excelvba5/EXCELVBA269.html


'Dictionaryにプロシージャー・プロパティ情報を格納
Public Sub GetCodeModule(ByRef dicProcInfo As Dictionary, _
                         ByVal wb As Workbook, _
                         ByVal sMod As String, _
                         ByVal ModuleType As String)
    Dim cProcInfo As ProcedureInfo
    Dim sProcName As String
    Dim sProcKey As String
    Dim iProcKind As Long: iProcKind = 0
    Dim i As Long
    Dim VVC As Object
    Set VVC = wb.VBProject.VBComponents(sMod).CodeModule
    i = 1
    Do While i <= VVC.CountOfLines
        sProcName = VVC.ProcOfLine(i, iProcKind)
        sProcKey = sMod & "." & sProcName
        If sProcName <> "" Then
            If IsProcLine(VVC.Lines(i, 1), sProcName) Then
                If Not dicProcInfo.Exists(sProcKey) Then
                    Set cProcInfo = New ProcedureInfo
                    cProcInfo.ModName = sMod
                    cProcInfo.ModType = ModuleType
                    cProcInfo.ProcName = sProcName
                    cProcInfo.ProcKind = iProcKind
                    cProcInfo.LineNo = i
                    cProcInfo.Comment = GetProcComment(i, VVC)
                    cProcInfo.Source = GetProcSource(i, VVC)
                    dicProcInfo.Add sProcKey, cProcInfo
                End If
            End If
        End If
        i = i + 1
    Loop
End Sub

'プロシージャー・プロパティ定義行かの判定
Private Function IsProcLine(ByVal strLine As String, _
                            ByVal ProcName As String) As Boolean
    strLine = " " & Trim(strLine)
    Select Case True
        Case Left(strLine, 1) = " '"
            IsProcLine = False
        Case strLine Like "* Sub " & ProcName & "(*"
            IsProcLine = True
        Case strLine Like "* Sub " & ProcName & " _"
            IsProcLine = True
        Case strLine Like "* Function " & ProcName & "(*"
            IsProcLine = True
        Case strLine Like "* Function " & ProcName & " _"
            IsProcLine = True
        Case strLine Like "* Property * " & ProcName & "(*"
            IsProcLine = True
        Case strLine Like "* Property * " & ProcName & " _"
            IsProcLine = True
        Case Else
            IsProcLine = False
    End Select
End Function

'継続行( _)全てを連結した文字列で返す
Private Function GetProcSource(ByRef i As Long, _
                               ByVal aCodeModule As Object) As String
    GetProcSource = ""
    Dim sTemp As String
    Do
        sTemp = Trim(aCodeModule.Lines(i, 1))
        If Right(aCodeModule.Lines(i, 1), 2) = " _" Then
            sTemp = Left(sTemp, Len(sTemp) - 1)
        End If
        GetProcSource = GetProcSource & sTemp
        If Right(aCodeModule.Lines(i, 1), 2) <> " _" Then Exit Do
        i = i + 1
    Loop
End Function

'プロシージャーの直前のコメントを取得
Private Function GetProcComment(ByVal i As Long, _
                                ByVal aCodeModule As Object) As String
    GetProcComment = ""
    i = i - 1
    Do While Left(aCodeModule.Lines(i, 1), 1) = "'"
        If GetProcComment <> "" Then GetProcComment = vbLf & GetProcComment
        GetProcComment = aCodeModule.Lines(i, 1) & GetProcComment
        i = i - 1
    Loop
End Function

