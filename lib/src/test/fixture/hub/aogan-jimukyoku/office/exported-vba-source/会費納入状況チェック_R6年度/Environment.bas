Attribute VB_Name = "Environment"
Option Explicit

Function mbGetPathOfAoganCashbook() As String
    ' このブックのなかの実行環境ワークシートの中に
    ' 青森県眼科医会の現金出納帳Excelブックのパスが書いてある。
    ' そのパスはこのExcelブックのパスを基底とする相対パスである。
    ' このパスの値をワークシートから読み出して、絶対パスに変換して、
    ' Stringとして返す。
    mbGetPathOfAoganCashbook = KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2")
End Function

Sub Test_mbGetPathOfAoganCashbook()
    ' GetPathOfAoganMembersをテストする
    Call KzCls
    'Sub Cls1は Cashbook現金出納帳のためのVBAライブラリ.xlam に含まれている
    Debug.Print mbGetPathOfAoganCashbook()
End Sub

Function mbGetPathOfAoganMembers() As String
    'ThisWorkbookのなかの「会員名簿ファイルのパス」ワークシートのなかに
    '青森県眼科医会の会員名簿Excelファイルのパスが書いてある。
    'そのパスはThisWorkbookのパスを基底とする相対パスである。
    'この値を読みだして絶対パスに変換してStringとして返す。
    mbGetPathOfAoganMembers = KzResolveExternalFilePath(ThisWorkbook, "会員名簿ファイルのパス", "B2")
End Function

Sub Test_mbGetPathOfAoganMembers()
    ' GetPathOfAoganMembersをテストする
    Call KzCls
    'Sub Cls1は Cashbook現金出納帳のためのVBAライブラリ.xlam に含まれている
    Debug.Print mbGetPathOfAoganMembers()
End Sub
