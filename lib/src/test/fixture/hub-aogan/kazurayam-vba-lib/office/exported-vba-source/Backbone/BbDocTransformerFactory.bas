Attribute VB_Name = "BbDocTransformerFactory"
Option Explicit

Public Function CreateDocTransformer() As BbDocTransformer
    Dim dt As BbDocTransformer: Set dt = New BbDocTransformer
    Set CreateDocTransformer = dt
End Function

