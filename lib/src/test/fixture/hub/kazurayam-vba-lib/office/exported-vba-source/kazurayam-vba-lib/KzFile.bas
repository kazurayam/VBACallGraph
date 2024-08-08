Attribute VB_Name = "KzFile"
Option Explicit

'KzFile

' givenPathが絶対パスならばTrue、さもなければFalseを返す
' "C:\something" は英字:\で始まっているから絶対パスと認める
' "\something"は \で始まっているから絶対パスと認める
' "..\something"は上記2つの条件に当てはまらないから相対パスとみなす
Public Function KzIsAbsolutePath(ByVal givenPath As String) As Boolean
    If Left(givenPath, 3) Like "[A-Z]:\" Or (Left(givenPath, 1) Like "\") Then
        KzIsAbsolutePath = True
    Else
        KzIsAbsolutePath = False
    End If
End Function

Public Function KzAbsolutifyPath(ByVal basePath As String, ByVal givenPath As String) As String
    ' 相対パスが与えられたら絶対パスに変換して返す。このときbasePathを基底として解釈する。
    ' 絶対パスが与えられたらそのまま返す
    If Not KzIsAbsolutePath(givenPath) Then
        Dim objFso As Object: Set objFso = CreateObject("Scripting.FileSystemObject")
        KzAbsolutifyPath = objFso.GetAbsolutePathName(objFso.BuildPath(basePath, givenPath))
        Set objFso = Nothing
    Else
        KzAbsolutifyPath = givenPath
    End If
End Function



' 引数pathが "https://d.docs.live.net/c5960fe753e170b9/デスクトップ/Excel-Word-VBA" のように
' そのファイルがOneDriveにマッピングされていることを示すURL文字列かどうかを調べる。
' もしそうならば "C:\Users" で始まるOneDriveのローカルな形式のStringに書きかえて返す。
' もしそうでなければpathをそのまま返す。
Public Function KzToLocalFilePath(ByVal path As String) As String
    Dim searchResult As Integer
    searchResult = VBA.Strings.InStr(1, path, "https://d.docs.live.net/", vbTextCompare)
    ' Debug.Print "searchResult=" & searchResult
    If searchResult = 1 Then
        Dim s() As String
        s = VBA.Strings.Split(path, "/", Limit:=5, Compare:=vbBinaryCompare)
        ' sは配列で中身は Arrays("https:", "", "d.docs.live.net", "c5960fe753e170b9", "デスクトップ/Excel-Word-VBA") になっている
        Dim objFso As Object
        Set objFso = CreateObject("Scripting.FileSystemObject")
        Dim p As String: p = objFso.GetAbsolutePathName(objFso.BuildPath(VBA.Interaction.Environ("OneDrive"), s(UBound(s))))
        ' UBound関数は引数に指定した配列で使用できる最も大きいインデックス番号を返す
        ' s(UBound(s)) は 配列sの5番目の要素 "デスクトップ/Excel-Word-VBA" を返す
        ' VBA.Interaction.Environ()は環境変数の値を返す
        ' Environ("OneDrive")の値はたとえば "C:\Users\uraya\OneDrive" という文字列を返す
        ' objFso.BuildPath(path, name)はパスとファイル名を表す二つの文字列を連結してひとつの文字列を返す。/は\に置き換えられる。
        ' objFso.GetAbsolutePathName(pathspec)は pathspec（相対パスかもしれない）を絶対パスに変換します
        KzToLocalFilePath = p
        Set objFso = Nothing
    Else
        KzToLocalFilePath = path
    End If
End Function


' Stringとしてパスが指定されたフォルダがすでに存在しているかどうかを調べて
' もしも未だ無かったら作る。ただし親フォルダも無い場合には失敗する。
Public Sub KzCreateFolder(folderPath As String)
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    If objFso.FolderExists(folderPath) Then
        ' does nothing
    Else
        objFso.CreateFolder (folderPath)
        Debug.Print "created " & folderPath
    End If
    Set objFso = Nothing
End Sub



' フォルダのフルパスが与えられることを前提する。フォルダを作る。
' ルートから子孫フォルダを順番に有無をしらべて、無ければMkDirで作る。
' つまり指定されたフォルダの先祖が無ければ先祖も作ってしまう。
Public Sub KzEnsureFolders(path As String)
    Dim tmp As String
    Dim arr() As String
    arr = Split(path, "\")
    tmp = arr(0)
    Dim i As Long
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = tmp & "\" & arr(i)
        If Dir(tmp, vbDirectory) = "" Then
            ' フォルダが無ければ作る
            MkDir tmp
        End If
    Next i
End Sub




' pathが示すパスにファイルまたはフォルダが存在していたらTrueをかえす。
' pathが示すパスにファイルもフォルダも無いならFalseをかえす
Public Function KzPathExists(ByVal path As String) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim flg As Boolean: flg = False
    If fso.FileExists(path) Then
        flg = True
    ElseIf fso.FolderExists(path) Then
        flg = True
    End If
    KzPathExists = flg
End Function





' パスを指定したファイルが存在していたら削除する。
' ファイルが無ければなにもしない。
Public Sub KzDeleteFile(ByVal fileToDelete As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(fileToDelete) Then 'See above
        ' First remove readonly attribute, if set
        SetAttr fileToDelete, vbNormal
        ' Then delete the file
        Kill fileToDelete
    End If
End Sub


' フォルダが存在していたら削除する
Public Sub KzDeleteFolder(ByVal folderToDelete As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderToDelete) Then
        fso.DeleteFolder (folderToDelete)
    End If
End Sub

' テキストをファイルにWRITEする。
' ファイルを納めるべき親フォルダが無ければ作ってから。
Public Sub KzWriteTextIntoFile(ByVal textData As String, ByVal file As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    KzEnsureFolders (fso.getParentFolderName(file))
    If fso.FileExists(file) Then
        KzDeleteFile (file)
    End If
    Dim fileNo As Integer
    fileNo = FreeFile
    Open file For Output As #fileNo
    Write #fileNo, textData
    Close #fileNo
End Sub



