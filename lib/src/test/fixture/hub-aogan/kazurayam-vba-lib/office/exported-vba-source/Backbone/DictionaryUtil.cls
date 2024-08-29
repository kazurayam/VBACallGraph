VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub PrintDictionary(dict As Dictionary)
    Dim key As Variant
    For Each key In dict.keys
        Debug.Print key, dict(key)
    Next key
End Sub



