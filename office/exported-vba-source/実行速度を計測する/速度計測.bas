Attribute VB_Name = "‘¬“xŒv‘ª"
Option Explicit

Sub measure()
                      
    Dim timerObj As TimerObject: Set timerObj = CreateTimer
    Dim booster As PerformanceBooster: Set booster = CreateBooster
    
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
