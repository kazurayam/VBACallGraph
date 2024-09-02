Attribute VB_Name = "BbLog"
Option Explicit

'BbLog

' Clear Immediate Window
' calls Debug.Print many times to output blank lines
' so that the immediate window is wiped out
Public Sub Clear()
    Debug.Print String(200, vbCrLf)
End Sub


' print a informational log in the format of "[moduleName.procesureName] message"
Public Sub Info(ByVal moduleName As String, ByVal procedureName As String, ByVal message As String)
    Debug.Print "[" & moduleName & "." & procedureName + "] " & message
End Sub

