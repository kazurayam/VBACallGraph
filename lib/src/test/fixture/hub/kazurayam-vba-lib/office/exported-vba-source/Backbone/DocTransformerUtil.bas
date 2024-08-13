Attribute VB_Name = "DocTransformerUtil"
Option Explicit

Public Function Create() As DocTransformer
    Dim DT As DocTransformer: Set DT = New DocTransformer
    Set Create = DT
End Function

