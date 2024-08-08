VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' メンバ変数
Public Name As String
Private Age As Long

' メソッド
Public Sub Initialize(ByVal nm As String, Optional ByVal ag As Long = 16)
    Name = nm
    Age = ag
End Sub

Public Sub SayHello()
    Debug.Print "Hello, I am " & Name & "!"
End Sub

Public Sub SayHelloTo(other As String)
    Debug.Print "Hello, " & other & "! I am " & Name & "!"
End Sub

Public Function GetAge() As Long
    GetAge = Age
End Function

Public Function ToString() As String
    ToString = Name
End Function

