Attribute VB_Name = "modMain"
'===============================================================================
' SDTD - Source to Target Document
' Module: modMain
' Purpose: Entry point, menu/toolbar creation, document selection
'===============================================================================
Option Explicit

' Global document references
Public gSourceDoc As Document
Public gTargetDoc As Document

' Settings stored in registry
Private Const REG_KEY As String = "SDTD"

'---------------------------------------------------------------
' AutoExec: Runs when the template loads (Word STARTUP)
'---------------------------------------------------------------
Public Sub AutoExec()
    CreateSDTDMenu
End Sub

'---------------------------------------------------------------
' AutoExit: Cleanup when Word closes
'---------------------------------------------------------------
Public Sub AutoExit()
    RemoveSDTDMenu
End Sub

'---------------------------------------------------------------
' Create the SDTD menu in the Add-ins tab
'---------------------------------------------------------------
Private Sub CreateSDTDMenu()
    Dim cb As CommandBar
    Dim cbPopup As CommandBarPopup
    Dim cbBtn As CommandBarButton
    
    On Error Resume Next
    ' Remove existing to avoid duplicates
    Application.CommandBars("Menu Bar").Controls("SDTD").Delete
    On Error GoTo 0
    
    Set cb = Application.CommandBars("Menu Bar")
    Set cbPopup = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbPopup.Caption = "SDTD"
    
    ' --- Show Task Pane ---
    Set cbBtn = cbPopup.Controls.Add(Type:=msoControlButton)
    cbBtn.Caption = "Open SDTD Panel"
    cbBtn.OnAction = "ShowSDTDPanel"
    cbBtn.FaceId = 2950
    
    ' --- Separator ---
    cbPopup.Controls.Add(Type:=msoControlButton).BeginGroup = True
    
    ' --- Layout ---
    Set cbBtn = cbPopup.Controls.Add(Type:=msoControlButton)
    cbBtn.Caption = "Copy Page Layout"
    cbBtn.OnAction = "RunCopyPageLayout"
    
    Set cbBtn = cbPopup.Controls.Add(Type:=msoControlButton)
    cbBtn.Caption = "Copy Styles"
    cbBtn.OnAction = "RunCopyStyles"
    
    ' --- Format ---
    Set cbBtn = cbPopup.Controls.Add(Type:=msoControlButton)
    cbBtn.Caption = "Format Helper"
    cbBtn.OnAction = "RunFormatHelper"
    cbBtn.BeginGroup = True
    
    Set cbBtn = cbPopup.Controls.Add(Type:=msoControlButton)
    cbBtn.Caption = "Clean the MEPS Up"
    cbBtn.OnAction = "RunCleanMEPS"
    
    ' --- ADP ---
    Set cbBtn = cbPopup.Controls.Add(Type:=msoControlButton)
    cbBtn.Caption = "Replace Key Phrases"
    cbBtn.OnAction = "RunReplaceKeyPhrases"
    cbBtn.BeginGroup = True
    
    Set cbBtn = cbPopup.Controls.Add(Type:=msoControlButton)
    cbBtn.Caption = "Transfer DP Comments"
    cbBtn.OnAction = "RunTransferDPComments"
    
    Set cbBtn = cbPopup.Controls.Add(Type:=msoControlButton)
    cbBtn.Caption = "Detect Missing Sections"
    cbBtn.OnAction = "RunDetectMissingSections"
    
    ' --- Batch ---
    Set cbBtn = cbPopup.Controls.Add(Type:=msoControlButton)
    cbBtn.Caption = "Batch Process..."
    cbBtn.OnAction = "RunBatchProcess"
    cbBtn.BeginGroup = True
End Sub

'---------------------------------------------------------------
' Remove the SDTD menu
'---------------------------------------------------------------
Private Sub RemoveSDTDMenu()
    On Error Resume Next
    Application.CommandBars("Menu Bar").Controls("SDTD").Delete
    On Error GoTo 0
End Sub

'---------------------------------------------------------------
' Show the SDTD UserForm (task pane)
'---------------------------------------------------------------
Public Sub ShowSDTDPanel()
    frmSDTD.Show vbModeless
End Sub

'---------------------------------------------------------------
' Select Source Document
'---------------------------------------------------------------
Public Function SelectSourceDocument() As Boolean
    Dim doc As Document
    Set doc = PickOpenDocument("Select SOURCE Document")
    If Not doc Is Nothing Then
        Set gSourceDoc = doc
        SelectSourceDocument = True
    End If
End Function

'---------------------------------------------------------------
' Select Target Document
'---------------------------------------------------------------
Public Function SelectTargetDocument() As Boolean
    Dim doc As Document
    Set doc = PickOpenDocument("Select TARGET Document")
    If Not doc Is Nothing Then
        Set gTargetDoc = doc
        SelectTargetDocument = True
    End If
End Function

'---------------------------------------------------------------
' Pick from currently open documents
'---------------------------------------------------------------
Private Function PickOpenDocument(title As String) As Document
    Dim docNames() As String
    Dim i As Integer
    Dim choice As String
    
    If Documents.Count = 0 Then
        MsgBox "No documents are open.", vbExclamation, "SDTD"
        Set PickOpenDocument = Nothing
        Exit Function
    End If
    
    ReDim docNames(1 To Documents.Count)
    Dim msg As String
    msg = title & vbCrLf & vbCrLf
    
    For i = 1 To Documents.Count
        docNames(i) = Documents(i).Name
        msg = msg & i & ". " & docNames(i) & vbCrLf
    Next i
    
    msg = msg & vbCrLf & "Enter number (1-" & Documents.Count & "):"
    choice = InputBox(msg, "SDTD - " & title)
    
    If choice = "" Then
        Set PickOpenDocument = Nothing
        Exit Function
    End If
    
    Dim idx As Integer
    idx = Val(choice)
    If idx >= 1 And idx <= Documents.Count Then
        Set PickOpenDocument = Documents(idx)
    Else
        MsgBox "Invalid selection.", vbExclamation, "SDTD"
        Set PickOpenDocument = Nothing
    End If
End Function

'---------------------------------------------------------------
' Check Paragraph Count for both documents
'---------------------------------------------------------------
Public Sub CheckParagraphCount()
    If Not ValidateDocs() Then Exit Sub
    
    Dim srcCount As Long, tgtCount As Long
    srcCount = gSourceDoc.Paragraphs.Count
    tgtCount = gTargetDoc.Paragraphs.Count
    
    Dim icon As VbMsgBoxStyle
    If srcCount = tgtCount Then
        icon = vbInformation
    Else
        icon = vbExclamation
    End If
    
    MsgBox "Source: " & srcCount & " paragraphs" & vbCrLf & _
           "Target: " & tgtCount & " paragraphs" & vbCrLf & vbCrLf & _
           IIf(srcCount = tgtCount, "✓ Paragraph counts match.", _
               "⚠ Mismatch! Difference: " & Abs(srcCount - tgtCount)), _
           icon, "SDTD - Paragraph Count"
End Sub

'---------------------------------------------------------------
' Validate that both documents are selected and open
'---------------------------------------------------------------
Public Function ValidateDocs() As Boolean
    If gSourceDoc Is Nothing Then
        MsgBox "Please select a Source document first.", vbExclamation, "SDTD"
        ValidateDocs = False
        Exit Function
    End If
    If gTargetDoc Is Nothing Then
        MsgBox "Please select a Target document first.", vbExclamation, "SDTD"
        ValidateDocs = False
        Exit Function
    End If
    
    ' Check if documents are still open
    Dim d As Document
    Dim srcOpen As Boolean, tgtOpen As Boolean
    For Each d In Documents
        If d.FullName = gSourceDoc.FullName Then srcOpen = True
        If d.FullName = gTargetDoc.FullName Then tgtOpen = True
    Next d
    
    If Not srcOpen Then
        MsgBox "Source document has been closed.", vbExclamation, "SDTD"
        Set gSourceDoc = Nothing
        ValidateDocs = False
        Exit Function
    End If
    If Not tgtOpen Then
        MsgBox "Target document has been closed.", vbExclamation, "SDTD"
        Set gTargetDoc = Nothing
        ValidateDocs = False
        Exit Function
    End If
    
    ValidateDocs = True
End Function

'---------------------------------------------------------------
' Save/Load settings from registry
'---------------------------------------------------------------
Public Sub SaveSetting_(key As String, value As String)
    SaveSetting "SDTD", "Settings", key, value
End Sub

Public Function GetSetting_(key As String, Optional default_ As String = "") As String
    GetSetting_ = GetSetting("SDTD", "Settings", key, default_)
End Function
