Attribute VB_Name = "modUtils"
'===============================================================================
' SDTD - Source to Target Document
' Module: modUtils
' Purpose: Shared utilities (progress bar, WorksheetFunction substitute)
'===============================================================================
Option Explicit

'---------------------------------------------------------------
' Progress display using status bar
'---------------------------------------------------------------
Public Sub ShowProgress(label As String, percent As Double)
    Application.StatusBar = "SDTD: " & label & " [" & String(Int(percent / 5), "|") & String(20 - Int(percent / 5), ".") & "] " & Int(percent) & "%"
    DoEvents
End Sub

Public Sub HideProgress()
    Application.StatusBar = False
End Sub

'---------------------------------------------------------------
' WorksheetFunction.Min substitute (no Excel dependency needed)
'---------------------------------------------------------------
Public Function MinVal(a As Long, b As Long) As Long
    If a < b Then MinVal = a Else MinVal = b
End Function

'---------------------------------------------------------------
' WorksheetFunction.Max substitute
'---------------------------------------------------------------
Public Function MaxVal(a As Long, b As Long) As Long
    If a > b Then MaxVal = a Else MaxVal = b
End Function
