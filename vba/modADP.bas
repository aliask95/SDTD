Attribute VB_Name = "modADP"
'===============================================================================
' SDTD - Source to Target Document
' Module: modADP
' Purpose: Replace Key Phrases, Transfer DP Comments, Detect Missing Sections
'===============================================================================
Option Explicit

'---------------------------------------------------------------
' Replace Key Phrases using Excel dictionary
' Column A = Find, Column B = Replace
'---------------------------------------------------------------
Public Sub RunReplaceKeyPhrases()
    ReplaceKeyPhrases False, False
End Sub

Public Sub ReplaceKeyPhrases(skipComments As Boolean, includeHeaderFooter As Boolean)
    If gTargetDoc Is Nothing Then
        MsgBox "Please select a Target document first.", vbExclamation, "SDTD"
        Exit Sub
    End If
    
    ' Get Excel file
    Dim excelPath As String
    excelPath = GetSetting_("LastExcelFile", "")
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select Excel Dictionary File"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
    If excelPath <> "" Then fd.InitialFileName = excelPath
    
    If fd.Show <> -1 Then Exit Sub
    excelPath = fd.SelectedItems(1)
    SaveSetting_ "LastExcelFile", excelPath
    
    ' Open Excel and read dictionary
    Dim xlApp As Object ' Excel.Application
    Dim xlWb As Object  ' Excel.Workbook
    Dim xlWs As Object  ' Excel.Worksheet
    
    On Error GoTo ExcelError
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWb = xlApp.Workbooks.Open(excelPath, ReadOnly:=True)
    Set xlWs = xlWb.Sheets(1)
    
    ' Read pairs into arrays
    Dim lastRow As Long
    lastRow = xlWs.Cells(xlWs.Rows.Count, 1).End(-4162).Row ' xlUp = -4162
    
    If lastRow < 1 Then
        MsgBox "Excel file appears empty.", vbExclamation, "SDTD"
        GoTo CleanupExcel
    End If
    
    Application.ScreenUpdating = False
    ShowProgress "Replacing key phrases...", 0
    
    Dim findText As String, replaceText As String
    Dim replacements As Long
    replacements = 0
    Dim r As Long
    
    For r = 1 To lastRow
        findText = CStr(xlWs.Cells(r, 1).value)
        replaceText = CStr(xlWs.Cells(r, 2).value)
        
        If Len(findText) > 0 Then
            ShowProgress "Replacing " & r & "/" & lastRow & ": " & Left(findText, 20), (r / lastRow) * 100
            
            ' Replace in main body
            Dim rng As Range
            Set rng = gTargetDoc.Content
            
            If skipComments Then
                ' Exclude comments by processing only the body range
                ' (Comments are separate story ranges)
                Set rng = gTargetDoc.StoryRanges(wdMainTextStory)
            End If
            
            replacements = replacements + FindAndReplaceInRange(rng, findText, replaceText)
            
            ' Replace in headers/footers if requested
            If includeHeaderFooter Then
                Dim sec As Section
                For Each sec In gTargetDoc.Sections
                    Dim hf As HeaderFooter
                    For Each hf In sec.Headers
                        If hf.Exists Then
                            replacements = replacements + FindAndReplaceInRange(hf.Range, findText, replaceText)
                        End If
                    Next hf
                    For Each hf In sec.Footers
                        If hf.Exists Then
                            replacements = replacements + FindAndReplaceInRange(hf.Range, findText, replaceText)
                        End If
                    Next hf
                Next sec
            End If
        End If
        
        DoEvents
    Next r
    
CleanupExcel:
    On Error Resume Next
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    On Error GoTo 0
    
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Replacement complete!" & vbCrLf & _
           replacements & " replacements made from " & lastRow & " dictionary entries.", _
           vbInformation, "SDTD"
    Exit Sub
    
ExcelError:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    On Error GoTo 0
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Error reading Excel file: " & Err.Description, vbCritical, "SDTD"
End Sub

'---------------------------------------------------------------
' Find and replace within a range, return count
'---------------------------------------------------------------
Private Function FindAndReplaceInRange(rng As Range, findText As String, replaceText As String) As Long
    Dim count As Long
    count = 0
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        
        Do While .Execute(Replace:=wdReplaceOne)
            count = count + 1
        Loop
    End With
    
    FindAndReplaceInRange = count
End Function

'---------------------------------------------------------------
' Transfer Digital Publishing Comments
' Copies DP comments from Source to corresponding Target paragraph
' Replaces identifier (e.g., "[ =E ]" -> "[ =BL ]")
' Optionally skips comments containing "it"
'---------------------------------------------------------------
Public Sub RunTransferDPComments()
    TransferDPComments "=E", "=BL", False
End Sub

Public Sub TransferDPComments(findId As String, replaceWith As String, excludeIt As Boolean)
    If Not ValidateDocs() Then Exit Sub
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ShowProgress "Transferring DP Comments...", 0
    
    Dim srcComment As Comment
    Dim transferred As Long, skipped As Long
    transferred = 0
    skipped = 0
    
    Dim total As Long
    total = gSourceDoc.Comments.Count
    
    If total = 0 Then
        MsgBox "No comments found in Source document.", vbInformation, "SDTD"
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To total
        Set srcComment = gSourceDoc.Comments(i)
        
        ShowProgress "Processing comment " & i & "/" & total, (i / total) * 100
        
        ' Only process "Digital Publishing" author
        If LCase(srcComment.Author) = LCase("Digital Publishing") Or _
           LCase(srcComment.Author) = "dp" Then
            
            ' Check exclusion
            If excludeIt And InStr(1, srcComment.Range.Text, "it", vbTextCompare) > 0 Then
                skipped = skipped + 1
                GoTo NextComment
            End If
            
            ' Find corresponding paragraph in target
            Dim srcParaIdx As Long
            srcParaIdx = GetParagraphIndex(gSourceDoc, srcComment.Scope)
            
            If srcParaIdx > 0 And srcParaIdx <= gTargetDoc.Paragraphs.Count Then
                ' Get comment text and replace identifier
                Dim commentText As String
                commentText = srcComment.Range.Text
                commentText = Replace(commentText, "[ " & findId & " ]", "[ " & replaceWith & " ]")
                commentText = Replace(commentText, "[" & findId & "]", "[" & replaceWith & "]")
                
                ' Get commented content from source
                Dim commentedContent As String
                commentedContent = srcComment.Scope.Text
                
                ' Append to end of target paragraph
                Dim tgtPara As Paragraph
                Set tgtPara = gTargetDoc.Paragraphs(srcParaIdx)
                
                Dim insertRange As Range
                Set insertRange = tgtPara.Range
                insertRange.Collapse wdCollapseEnd
                insertRange.MoveEnd wdCharacter, -1 ' Before paragraph mark
                
                ' Insert the commented content with the comment
                insertRange.InsertAfter " " & commentedContent
                
                ' Add comment to the inserted text
                Dim newRange As Range
                Set newRange = gTargetDoc.Range( _
                    insertRange.Start, _
                    insertRange.Start + Len(commentedContent) + 1)
                
                Dim newComment As Comment
                Set newComment = gTargetDoc.Comments.Add(newRange, commentText)
                newComment.Author = "Digital Publishing"
                newComment.Initial = "DP"
                
                transferred = transferred + 1
            Else
                skipped = skipped + 1
            End If
        End If
        
NextComment:
        DoEvents
    Next i
    
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Transfer complete!" & vbCrLf & _
           transferred & " comments transferred." & vbCrLf & _
           skipped & " comments skipped.", vbInformation, "SDTD"
    Exit Sub
    
ErrorHandler:
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Error transferring comments: " & Err.Description, vbCritical, "SDTD"
End Sub

'---------------------------------------------------------------
' Get the paragraph index of a range
'---------------------------------------------------------------
Private Function GetParagraphIndex(doc As Document, rng As Range) As Long
    Dim i As Long
    For i = 1 To doc.Paragraphs.Count
        If rng.Start >= doc.Paragraphs(i).Range.Start And _
           rng.Start < doc.Paragraphs(i).Range.End Then
            GetParagraphIndex = i
            Exit Function
        End If
    Next i
    GetParagraphIndex = 0
End Function

'---------------------------------------------------------------
' Detect Missing Sections
' Compares headings in Source vs Target
'---------------------------------------------------------------
Public Sub RunDetectMissingSections()
    If Not ValidateDocs() Then Exit Sub
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ShowProgress "Detecting missing sections...", 0
    
    ' Collect headings from Source
    Dim srcHeadings As New Collection ' Text -> heading level
    Dim tgtHeadings As New Collection
    
    Dim p As Paragraph
    Dim headingText As String
    Dim total As Long, current As Long
    
    ' Source headings
    total = gSourceDoc.Paragraphs.Count
    current = 0
    For Each p In gSourceDoc.Paragraphs
        current = current + 1
        If current Mod 100 = 0 Then ShowProgress "Scanning source... " & current & "/" & total, (current / total) * 50
        
        If IsHeadingStyle(p.Style) Then
            headingText = Trim(Replace(p.Range.Text, Chr(13), ""))
            If Len(headingText) > 0 Then
                On Error Resume Next
                srcHeadings.Add headingText, headingText
                On Error GoTo ErrorHandler
            End If
        End If
    Next p
    
    ' Target headings
    total = gTargetDoc.Paragraphs.Count
    current = 0
    For Each p In gTargetDoc.Paragraphs
        current = current + 1
        If current Mod 100 = 0 Then ShowProgress "Scanning target... " & current & "/" & total, 50 + (current / total) * 50
        
        If IsHeadingStyle(p.Style) Then
            headingText = Trim(Replace(p.Range.Text, Chr(13), ""))
            If Len(headingText) > 0 Then
                On Error Resume Next
                tgtHeadings.Add headingText, headingText
                On Error GoTo ErrorHandler
            End If
        End If
    Next p
    
    ' Find missing
    Dim missing As String
    Dim missingCount As Long
    missingCount = 0
    
    Dim idx As Long
    For idx = 1 To srcHeadings.Count
        Dim srcH As String
        srcH = srcHeadings(idx)
        
        ' Check if exists in target
        Dim found As Boolean
        found = False
        On Error Resume Next
        Dim test As String
        test = tgtHeadings(srcH)
        found = (Err.Number = 0)
        On Error GoTo ErrorHandler
        
        If Not found Then
            missingCount = missingCount + 1
            missing = missing & missingCount & ". " & srcH & vbCrLf
        End If
    Next idx
    
    HideProgress
    Application.ScreenUpdating = True
    
    If missingCount = 0 Then
        MsgBox "All source headings found in target!", vbInformation, "SDTD"
    Else
        MsgBox "Found " & missingCount & " missing section(s):" & vbCrLf & vbCrLf & _
               Left(missing, 1000), vbExclamation, "SDTD - Missing Sections"
    End If
    Exit Sub
    
ErrorHandler:
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Error detecting sections: " & Err.Description, vbCritical, "SDTD"
End Sub

'---------------------------------------------------------------
' Check if a style is a heading style
'---------------------------------------------------------------
Private Function IsHeadingStyle(sty As Variant) As Boolean
    On Error Resume Next
    Dim styleName As String
    styleName = LCase(CStr(sty))
    IsHeadingStyle = (Left(styleName, 7) = "heading" Or _
                      InStr(styleName, "heading") > 0 Or _
                      sty = wdStyleHeading1 Or _
                      sty = wdStyleHeading2 Or _
                      sty = wdStyleHeading3 Or _
                      sty = wdStyleHeading4 Or _
                      sty = wdStyleHeading5)
    On Error GoTo 0
End Function
