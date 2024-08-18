VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Accountクラス

' 勘定科目の名称を保持するオブジェクト

Enum AccountType
    Expense = 1
    Income = 2
    Undefined = 0
End Enum

Private accountType_ As Long
Private accountName_ As String
Private subAccountName_ As String

Public Sub Initialize(ByVal aType As Long, ByVal aName As String, ByVal sName As String)
    accountType_ = aType
    accountName_ = aName
    subAccountName_ = sName
End Sub

Public Function ToString() As String
    ToString = AccountTypeAsString & "/" & accountName_ & "/" & subAccountName_
End Function

Public Property Get accType() As Long
    accType = accountType_
End Property


Public Property Get AccountTypeAsString() As String
    Dim typeStr As String
    If accountType_ = AccountType.Expense Then
        typeStr = "支出"
    ElseIf accountType_ = AccountType.Income Then
        typeStr = "収入"
    Else
        typeStr = "0"
    End If
    AccountTypeAsString = typeStr
End Property

' Stringを引数とするイニシャライザ
' "支出/事務費/通信費"のような文字列を引数として受け取る
' 受け取った文字列を「支出」「事務費」「通信費」のように部分に分けて
' それに基づいてAccountオブジェクトのプロパティを初期化する。
Public Sub of(ByVal accStr As String)
    Dim arr() As String
    arr = Split(accStr, "/")
    'Debug.Print "UBound(arr)=" & UBound(arr)
    'Debug.Print "LBound(arr)=" & LBound(arr)
    If (UBound(arr) - LBound(arr)) = 2 Then
        If arr(0) Like "支出" Then
            accountType_ = AccountType.Expense
        ElseIf arr(0) Like "収入" Then
            accountType_ = AccountType.Income
        Else
            accountType_ = AccountType.Undefined
        End If
        accountName_ = arr(1)
        subAccountName_ = arr(2)
    Else
        Debug.Print "unable to parse addStr=" & accStr
    End If
End Sub

Public Property Get AccountName() As String
    AccountName = accountName_
End Property

Public Property Get SubAccountName() As String
    SubAccountName = subAccountName_
End Property




