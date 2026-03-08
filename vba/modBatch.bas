Attribute VB_Name = "modBatch"
'===============================================================================
' SDTD - Source to Target Document
' Module: modBatch
' Purpose: Batch processing across multiple files
'===============================================================================
Option Explicit

'---------------------------------------------------------------
' Batch Process: Apply selected operations to multiple files
'---------------------------------------------------------------
Public Sub RunBatchProcess()
    If gSourceDoc Is Nothing Then
        MsgBox "Please select a Source document first. It will be used as the " & _
               "layout/style template for all target files.", vbExclamation, "SDTD"
        Exit Sub
    End If
    
    ' Select target files
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select Target Files for Batch Processing"
    fd.Filters.Clear
    fd.Filters.Add "Word Documents", "*.docx;*.doc;*.docm"
    fd.AllowMultiSelect = True
    
    If fd.Show <> -1 Then Exit Sub
    
    Dim fileCount As Long
    fileCount = fd.SelectedItems.Count
    
    If fileCount = 0 Then Exit Sub
    
    ' Ask what operations to perform
    Dim ops As String
    ops = InputBox( _
        "Select operations to perform (enter numbers separated by commas):" & vbCrLf & vbCrLf & _
        "1. Copy Page Layout" & vbCrLf & _
        "2. Copy Styles" & vbCrLf & _
        "3. Format Helper" & vbCrLf & _
        "4. Clean MEPS (Bookman WTS + Universal)" & vbCrLf & _
        "5. Replace Key Phrases" & vbCrLf & vbCrLf & _
        "Example: 1,2,4", _
        "SDTD - Batch Operations", "1,2")
    
    If ops = "" Then Exit Sub
    
    Dim doLayout As Boolean, doStyles As Boolean, doFormat As Boolean
    Dim doMEPS As Boolean, doReplace As Boolean
    
    doLayout = InStr(ops, "1") > 0
    doStyles = InStr(ops, "2") > 0
    doFormat = InStr(ops, "3") > 0
    doMEPS = InStr(ops, "4") > 0
    doReplace = InStr(ops, "5") > 0
    
    ' Confirm
    Dim confirmMsg As String
    confirmMsg = "Batch processing " & fileCount & " file(s) with:" & vbCrLf
    If doLayout Then confirmMsg = confirmMsg & "  ✓ Copy Page Layout" & vbCrLf
    If doStyles Then confirmMsg = confirmMsg & "  ✓ Copy Styles" & vbCrLf
    If doFormat Then confirmMsg = confirmMsg & "  ✓ Format Helper" & vbCrLf
    If doMEPS Then confirmMsg = confirmMsg & "  ✓ Clean MEPS" & vbCrLf
    If doReplace Then confirmMsg = confirmMsg & "  ✓ Replace Key Phrases" & vbCrLf
    confirmMsg = confirmMsg & vbCrLf & "Continue?"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "SDTD - Confirm Batch") <> vbYes Then Exit Sub
    
    ' Process each file
    Application.ScreenUpdating = False
    
    Dim i As Long
    Dim processed As Long, errors As Long
    Dim errorLog As String
    processed = 0
    errors = 0
    
    For i = 1 To fileCount
        Dim filePath As String
        filePath = fd.SelectedItems(i)
        
        ShowProgress "Processing " & i & "/" & fileCount & ": " & Dir(filePath), (i / fileCount) * 100
        
        On Error Resume Next
        
        ' Open target document
        Dim tgtDoc As Document
        Set tgtDoc = Documents.Open(filePath, ReadOnly:=False, Visible:=False)
        
        If Err.Number <> 0 Then
            errors = errors + 1
            errorLog = errorLog & "Failed to open: " & Dir(filePath) & " - " & Err.Description & vbCrLf
            Err.Clear
            GoTo NextFile
        End If
        
        ' Set as global target
        Set gTargetDoc = tgtDoc
        
        ' Run selected operations
        If doLayout Then CopyPageLayout True, True, True, False, False
        If doStyles Then RunCopyStyles
        If doFormat Then RunFormatHelper
        If doMEPS Then CleanMEPS True, True
        If doReplace Then ReplaceKeyPhrases False, False
        
        ' Save and close
        tgtDoc.Save
        tgtDoc.Close
        Set tgtDoc = Nothing
        
        If Err.Number <> 0 Then
            errors = errors + 1
            errorLog = errorLog & "Error processing: " & Dir(filePath) & " - " & Err.Description & vbCrLf
            Err.Clear
        Else
            processed = processed + 1
        End If
        
NextFile:
        On Error GoTo 0
        DoEvents
    Next i
    
    Set gTargetDoc = Nothing
    HideProgress
    Application.ScreenUpdating = True
    
    Dim resultMsg As String
    resultMsg = "Batch processing complete!" & vbCrLf & _
                processed & " file(s) processed successfully." & vbCrLf & _
                errors & " error(s)."
    
    If errors > 0 Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & "Errors:" & vbCrLf & Left(errorLog, 500)
    End If
    
    MsgBox resultMsg, IIf(errors > 0, vbExclamation, vbInformation), "SDTD - Batch Results"
End Sub
