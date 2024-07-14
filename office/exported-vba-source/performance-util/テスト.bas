Attribute VB_Name = "ƒeƒXƒg"
Option Explicit

Sub MySub()
                      
    Dim timerObj As TimerObject: Set timerObj = New TimerObject
    Dim booster As PerformanceBooster: Set booster = New PerformanceBooster
    
    Sheet1.Cells.Clear
    Sheet2.Cells.Clear
    
    With Sheet1
        Dim i As Long
        For i = 1 To 300
            .Cells(i, 1).Value = i
            .Cells(i, 2).FormulaLocal = "=SUM(A1:A" & i & ")"
            .Rows(i).Copy
            Sheet2.Cells(i, 1).PasteSpecial
        Next i
    End With
                    
    timerObj.ReportTimer
    
End Sub
