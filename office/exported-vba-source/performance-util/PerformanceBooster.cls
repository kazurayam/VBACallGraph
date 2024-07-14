VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceBooster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private initCalculationValue_ As XlCalculation


Private Sub Class_Initialize()
    With Application
        initCalculationValue_ = .Calculation
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With
End Sub


Private Sub Class_Terminate()
    With Application
        .Calculation = initCalculationValue_
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub
