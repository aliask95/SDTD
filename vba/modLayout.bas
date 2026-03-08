Attribute VB_Name = "modLayout"
'===============================================================================
' SDTD - Source to Target Document
' Module: modLayout
' Purpose: Copy Page Layout, Copy Styles
'===============================================================================
Option Explicit

'---------------------------------------------------------------
' Copy Page Layout from Source to Target
'---------------------------------------------------------------
Public Sub RunCopyPageLayout()
    CopyPageLayout True, True, True, False, False
End Sub

Public Sub CopyPageLayout( _
    copyPageSize As Boolean, _
    copyHeader As Boolean, _
    copyFooter As Boolean, _
    copyHeaderFooterContent As Boolean, _
    deleteEmptyParas As Boolean)
    
    If Not ValidateDocs() Then Exit Sub
    
    Dim srcSec As Section, tgtSec As Section
    Dim srcPS As PageSetup, tgtPS As PageSetup
    Dim i As Long
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ShowProgress "Applying Layout...", 0
    
    ' Process each section (match by index)
    Dim secCount As Long
    secCount = WorksheetFunction.Min(gSourceDoc.Sections.Count, gTargetDoc.Sections.Count)
    
    For i = 1 To secCount
        Set srcSec = gSourceDoc.Sections(i)
        Set tgtSec = gTargetDoc.Sections(i)
        Set srcPS = srcSec.PageSetup
        Set tgtPS = tgtSec.PageSetup
        
        ShowProgress "Processing section " & i & "/" & secCount, (i / secCount) * 100
        
        ' --- Page Size ---
        If copyPageSize Then
            tgtPS.PageWidth = srcPS.PageWidth
            tgtPS.PageHeight = srcPS.PageHeight
            tgtPS.Orientation = srcPS.Orientation
            tgtPS.TopMargin = srcPS.TopMargin
            tgtPS.BottomMargin = srcPS.BottomMargin
            tgtPS.LeftMargin = srcPS.LeftMargin
            tgtPS.RightMargin = srcPS.RightMargin
            tgtPS.Gutter = srcPS.Gutter
            tgtPS.GutterPos = srcPS.GutterPos
            tgtPS.MirrorMargins = srcPS.MirrorMargins
            tgtPS.HeaderDistance = srcPS.HeaderDistance
            tgtPS.FooterDistance = srcPS.FooterDistance
            tgtPS.SectionStart = srcPS.SectionStart
        End If
        
        ' --- Header ---
        If copyHeader Then
            CopyHeaderFooterSettings srcSec, tgtSec, True, copyHeaderFooterContent, deleteEmptyParas
        End If
        
        ' --- Footer ---
        If copyFooter Then
            CopyHeaderFooterSettings srcSec, tgtSec, False, copyHeaderFooterContent, deleteEmptyParas
        End If
    Next i
    
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Page layout copied successfully!", vbInformation, "SDTD"
    Exit Sub
    
ErrorHandler:
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Error copying layout: " & Err.Description, vbCritical, "SDTD"
End Sub

'---------------------------------------------------------------
' Copy Header/Footer settings between sections
'---------------------------------------------------------------
Private Sub CopyHeaderFooterSettings( _
    srcSec As Section, _
    tgtSec As Section, _
    isHeader As Boolean, _
    copyContent As Boolean, _
    deleteEmptyParas As Boolean)
    
    Dim srcHF As HeaderFooter, tgtHF As HeaderFooter
    Dim hfTypes As Variant
    hfTypes = Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)
    
    Dim idx As Long
    For idx = LBound(hfTypes) To UBound(hfTypes)
        If isHeader Then
            Set srcHF = srcSec.Headers(hfTypes(idx))
            Set tgtHF = tgtSec.Headers(hfTypes(idx))
        Else
            Set srcHF = srcSec.Footers(hfTypes(idx))
            Set tgtHF = tgtSec.Footers(hfTypes(idx))
        End If
        
        ' Link to previous
        tgtHF.LinkToPrevious = srcHF.LinkToPrevious
        
        ' Copy content if requested
        If copyContent And srcHF.Exists Then
            srcHF.Range.Copy
            tgtHF.Range.PasteAndFormat wdFormatOriginalFormatting
            
            ' Delete empty paragraphs if requested
            If deleteEmptyParas Then
                DeleteEmptyParagraphs tgtHF.Range
            End If
        End If
    Next idx
    
    ' Copy "Different first page" and "Different odd & even" settings
    Dim srcPS As PageSetup, tgtPS As PageSetup
    Set srcPS = srcSec.PageSetup
    Set tgtPS = tgtSec.PageSetup
    tgtPS.DifferentFirstPageHeaderFooter = srcPS.DifferentFirstPageHeaderFooter
    tgtPS.OddAndEvenPagesHeaderFooter = srcPS.OddAndEvenPagesHeaderFooter
End Sub

'---------------------------------------------------------------
' Delete empty paragraphs in a range
'---------------------------------------------------------------
Private Sub DeleteEmptyParagraphs(rng As Range)
    Dim p As Paragraph
    Dim i As Long
    For i = rng.Paragraphs.Count To 1 Step -1
        Set p = rng.Paragraphs(i)
        If Len(Trim(p.Range.Text)) <= 1 Then  ' paragraph mark only
            If rng.Paragraphs.Count > 1 Then  ' keep at least one
                p.Range.Delete
            End If
        End If
    Next i
End Sub

'---------------------------------------------------------------
' Copy Styles from Source to Target
'---------------------------------------------------------------
Public Sub RunCopyStyles()
    If Not ValidateDocs() Then Exit Sub
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ShowProgress "Copying Styles...", 0
    
    Dim srcStyle As Style
    Dim i As Long, total As Long
    total = gSourceDoc.Styles.Count
    
    For i = 1 To total
        Set srcStyle = gSourceDoc.Styles(i)
        
        ' Only copy user-modified or custom styles
        If Not srcStyle.BuiltIn Or srcStyle.InUse Then
            On Error Resume Next
            gTargetDoc.Styles.Add srcStyle.NameLocal, srcStyle.Type
            On Error GoTo ErrorHandler
            
            ' Copy style properties
            On Error Resume Next
            gTargetDoc.Styles(srcStyle.NameLocal).Font.Name = srcStyle.Font.Name
            gTargetDoc.Styles(srcStyle.NameLocal).Font.Size = srcStyle.Font.Size
            gTargetDoc.Styles(srcStyle.NameLocal).Font.Bold = srcStyle.Font.Bold
            gTargetDoc.Styles(srcStyle.NameLocal).Font.Italic = srcStyle.Font.Italic
            gTargetDoc.Styles(srcStyle.NameLocal).Font.Color = srcStyle.Font.Color
            gTargetDoc.Styles(srcStyle.NameLocal).ParagraphFormat.Alignment = srcStyle.ParagraphFormat.Alignment
            gTargetDoc.Styles(srcStyle.NameLocal).ParagraphFormat.SpaceBefore = srcStyle.ParagraphFormat.SpaceBefore
            gTargetDoc.Styles(srcStyle.NameLocal).ParagraphFormat.SpaceAfter = srcStyle.ParagraphFormat.SpaceAfter
            gTargetDoc.Styles(srcStyle.NameLocal).ParagraphFormat.LineSpacingRule = srcStyle.ParagraphFormat.LineSpacingRule
            gTargetDoc.Styles(srcStyle.NameLocal).ParagraphFormat.LineSpacing = srcStyle.ParagraphFormat.LineSpacing
            On Error GoTo ErrorHandler
        End If
        
        If i Mod 20 = 0 Then
            ShowProgress "Copying styles... " & i & "/" & total, (i / total) * 100
        End If
    Next i
    
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Styles copied successfully!", vbInformation, "SDTD"
    Exit Sub
    
ErrorHandler:
    HideProgress
    Application.ScreenUpdating = True
    MsgBox "Error copying styles: " & Err.Description, vbCritical, "SDTD"
End Sub
