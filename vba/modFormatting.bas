Attribute VB_Name = "modFormatting"
'===============================================================================
' SDTD - Source to Target Document
' Module: modFormatting
' Purpose: Format Helper, Clean the MEPS Up
'===============================================================================
Option Explicit

'---------------------------------------------------------------
' Format Helper
' Detects Italic, Bold, Bold+Italic in Source and adds review
' comments in Target. Comment author: "Format Assist"
'---------------------------------------------------------------
Public Sub RunFormatHelper()
    If Not ValidateDocs() Then Exit Sub
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ShowProgress "Running Format Helper...", 0
    
    Dim srcParas As Paragraphs, tgtParas As Paragraphs
    Set srcParas = gSourceDoc.Paragraphs
    Set tgtParas = gTargetDoc.Paragraphs
    
    Dim paraCount As Long
    paraCount = WorksheetFunction.Min(srcParas.Count, tgtParas.Count)
    
    Dim i As Long
    Dim commentsAdded As Long
    commentsAdded = 0
    
    For i = 1 To paraCount
        If i Mod 50 = 0 Then
            ShowProgress "Scanning paragraph " & i & "/" & paraCount, (i / paraCount) * 100
            DoEvents
        End If
        
        Dim srcPara As Paragraph
        Set srcPara = srcParas(i)
        
        ' Skip empty paragraphs
        If Len(Trim(Replace(srcPara.Range.Text, Chr(13), ""))) = 0 Then GoTo NextPara
        
        ' Scan each word/run for formatting
        Dim srcRange As Range
        Set srcRange = srcPara.Range
        
        Dim charIdx As Long
        Dim runStart As Long, runEnd As Long
        Dim currentFormat As String, prevFormat As String
        
        prevFormat = ""
        runStart = srcRange.Start
        
        For charIdx = 1 To srcRange.Characters.Count
            Dim c As Range
            Set c = srcRange.Characters(charIdx)
            
            ' Determine format
            currentFormat = ""
            If c.Font.Bold And c.Font.Italic Then
                currentFormat = "Bold+Italic"
            ElseIf c.Font.Bold Then
                currentFormat = "Bold"
            ElseIf c.Font.Italic Then
                currentFormat = "Italic"
            End If
            
            ' When format changes, record the previous run
            If currentFormat <> prevFormat Then
                If prevFormat <> "" And Len(prevFormat) > 0 Then
                    ' Add comment to target at corresponding paragraph
                    Dim commentText As String
                    Dim formattedText As String
                    
                    Dim runRange As Range
                    Set runRange = gSourceDoc.Range(runStart, srcRange.Characters(charIdx - 1).End)
                    formattedText = Left(Trim(runRange.Text), 50)
                    
                    If Len(formattedText) > 0 And i <= tgtParas.Count Then
                        commentText = "[" & prevFormat & "] """ & formattedText & """"
                        
                        Dim tgtRange As Range
                        Set tgtRange = tgtParas(i).Range
                        
                        Dim newComment As Comment
                        Set newComment = gTargetDoc.Comments.Add(tgtRange, commentText)
                        newComment.Author = "Format Assist"
                        newComment.Initial = "FA"
                        commentsAdded = commentsAdded + 1
                    End If
                End If
                runStart = c.Start
                prevFormat = currentFormat
            End If
        Next charIdx
        
        ' Handle last run
        If prevFormat <> "" And i <= tgtParas.Count Then
            Dim lastRange As Range
            Set lastRange = gSourceDoc.Range(runStart, srcRange.End - 1)
            formattedText = Left(Trim(lastRange.Text), 50)
            
            If Len(formattedText) > 0 Then
                commentText = "[" & prevFormat & "] """ & formattedText & """"
                Set tgtRange = tgtParas(i).Range
                Set newComment = gTargetDoc.Comments.Add(tgtRange, commentText)
                newComment.Author = "Format Assist"
                newComment.Initial = "FA"
                commentsAdded = commentsAdded + 1
            End If
        End If
        
NextPara:
    Next i
    
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Format Helper complete!" & vbCrLf & _
           commentsAdded & " formatting comments added.", vbInformation, "SDTD"
    Exit Sub
    
ErrorHandler:
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Error in Format Helper: " & Err.Description, vbCritical, "SDTD"
End Sub

'---------------------------------------------------------------
' Clean the MEPS Up
' Removes text with MEPS fonts from Target document
'---------------------------------------------------------------
Public Sub RunCleanMEPS()
    CleanMEPS True, True
End Sub

Public Sub CleanMEPS(cleanBookmanWTS As Boolean, cleanBookmanUniversal As Boolean)
    If gTargetDoc Is Nothing Then
        MsgBox "Please select a Target document first.", vbExclamation, "SDTD"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ShowProgress "Cleaning MEPS fonts...", 0
    
    Dim fontsToClean As New Collection
    If cleanBookmanWTS Then fontsToClean.Add "MEPS Bookman WTS"
    If cleanBookmanUniversal Then fontsToClean.Add "MEPS Bookman Universal"
    
    If fontsToClean.Count = 0 Then
        MsgBox "No MEPS fonts selected.", vbExclamation, "SDTD"
        Exit Sub
    End If
    
    Dim totalRemoved As Long
    totalRemoved = 0
    
    Dim fontName As Variant
    Dim fIdx As Long
    fIdx = 0
    
    For Each fontName In fontsToClean
        fIdx = fIdx + 1
        ShowProgress "Removing " & fontName & "...", (fIdx / fontsToClean.Count) * 100
        
        ' Use Find & Replace to remove text with this font
        With gTargetDoc.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ""
            .Font.Name = CStr(fontName)
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            
            Do While .Execute(Replace:=wdReplaceOne)
                totalRemoved = totalRemoved + 1
                DoEvents
            Loop
        End With
    Next fontName
    
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "MEPS cleanup complete!" & vbCrLf & _
           totalRemoved & " instances removed.", vbInformation, "SDTD"
    Exit Sub
    
ErrorHandler:
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Error cleaning MEPS: " & Err.Description, vbCritical, "SDTD"
End Sub
