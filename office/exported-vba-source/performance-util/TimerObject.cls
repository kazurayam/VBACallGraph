VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Start As Date
Public Finish As Date

Private Sub Class_Initialize()
    Start = Time
End Sub

Public Sub ReportTimer()
    Finish = Time
    MsgBox "é¿çséûä‘ÇÕ " & Format(Finish - Start, "nnï™ssïb") & " Ç≈ÇµÇΩ", vbInformation + vbOKOnly
End Sub

