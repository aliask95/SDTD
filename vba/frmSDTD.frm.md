# SDTD UserForm — Design Instructions

Since UserForms can't be created as text files, follow these steps in the VBA Editor:

## Create the UserForm

1. **Insert → UserForm** — name it `frmSDTD`
2. Set properties:
   - **Caption**: `SDTD - Source to Target Document`
   - **Width**: 320
   - **Height**: 520
   - **BackColor**: &H00362B25& (dark background)
   - **StartUpPosition**: 0 - Manual

## Add Controls (top to bottom)

### Header Area (top)
- **Label** `lblTitle`: Caption = "SDTD", Font = Tahoma 14 Bold, ForeColor = White
- **Label** `lblSubtitle`: Caption = "Source → Target Document", Font = Tahoma 8, ForeColor = Gray

### Document Selection
- **Label** `lblSource`: Caption = "SOURCE"
- **ComboBox** `cboSource`: Lists open documents, Width = 200
- **CommandButton** `btnPickSource`: Caption = "...", Width = 24
- **Label** `lblTarget`: Caption = "TARGET"
- **ComboBox** `cboTarget`: Lists open documents, Width = 200
- **CommandButton** `btnPickTarget`: Caption = "...", Width = 24
- **CommandButton** `btnCheckParas`: Caption = "CHECK PARAGRAPH COUNT", Width = 280

### Tab Strip
- **TabStrip** `tabMain`: 3 tabs — "Layout", "Format", "ADP"
- **Width**: 290, **Height**: 280

### Layout Tab (Index 0) — inside a Frame `fraLayout`
- **CheckBox** `chkPageSize`: Caption = "Page Size", Value = True
- **CheckBox** `chkHeader`: Caption = "Header", Value = True
- **CheckBox** `chkFooter`: Caption = "Footer", Value = True
- **CheckBox** `chkCopyHFContent`: Caption = "Copy Header/Footer Content"
- **CheckBox** `chkDeleteEmptyParas`: Caption = "Delete empty paragraphs in header/footer" (indented, visible when chkCopyHFContent = True)
- **CommandButton** `btnApplyLayout`: Caption = "Apply Layout"
- **CommandButton** `btnCopyStyles`: Caption = "Copy Styles"

### Format Tab (Index 1) — inside a Frame `fraFormat`
- **CommandButton** `btnFormatHelper`: Caption = "Run Format Helper"
- **CheckBox** `chkMEPSWTS`: Caption = "MEPS Bookman WTS", Value = True
- **CheckBox** `chkMEPSUniversal`: Caption = "MEPS Bookman Universal", Value = True
- **CommandButton** `btnCleanMEPS`: Caption = "Clean the MEPS Up"

### ADP Tab (Index 2) — inside a Frame `fraADP`
- **CommandButton** `btnSelectExcel`: Caption = "Select Excel File"
- **Label** `lblExcelFile`: Caption = "(no file selected)"
- **CheckBox** `chkSkipComments`: Caption = "Exclude comments from replacement"
- **CheckBox** `chkIncludeHF`: Caption = "Include header and footer"
- **CommandButton** `btnRunReplace`: Caption = "Run Replacement"
- **TextBox** `txtFindId`: Text = "=E", Width = 80
- **TextBox** `txtReplaceWith`: Text = "=BL", Width = 80
- **CheckBox** `chkExcludeIt`: Caption = 'Exclude comments containing "it"'
- **CommandButton** `btnTransferDP`: Caption = "Transfer Comments"
- **CommandButton** `btnDetectMissing`: Caption = "Run Detection"

### Bottom Area
- **CheckBox** `chkBatchMode`: Caption = "Enable Batch Mode"
- **CommandButton** `btnSelectFiles`: Caption = "Select Files" (visible when batch mode on)
- **CommandButton** `btnSelectFolder`: Caption = "Select Folder" (visible when batch mode on)

---

## UserForm Code (paste into the form's code module)

```vb
Option Explicit

Private Sub UserForm_Initialize()
    RefreshDocumentLists
    tabMain.Value = 0
    ShowTab 0
    
    ' Load saved Excel file
    lblExcelFile.Caption = GetSetting_("LastExcelFile", "(no file selected)")
    
    ' Batch controls hidden initially
    btnSelectFiles.Visible = False
    btnSelectFolder.Visible = False
    chkDeleteEmptyParas.Visible = False
End Sub

Private Sub UserForm_Activate()
    RefreshDocumentLists
End Sub

'--- Document Selection ---
Private Sub RefreshDocumentLists()
    cboSource.Clear
    cboTarget.Clear
    Dim d As Document
    For Each d In Documents
        cboSource.AddItem d.Name
        cboTarget.AddItem d.Name
    Next d
    
    If Not gSourceDoc Is Nothing Then
        cboSource.Value = gSourceDoc.Name
    End If
    If Not gTargetDoc Is Nothing Then
        cboTarget.Value = gTargetDoc.Name
    End If
End Sub

Private Sub cboSource_Change()
    On Error Resume Next
    Set gSourceDoc = Documents(cboSource.Value)
    On Error GoTo 0
End Sub

Private Sub cboTarget_Change()
    On Error Resume Next
    Set gTargetDoc = Documents(cboTarget.Value)
    On Error GoTo 0
End Sub

Private Sub btnPickSource_Click()
    SelectSourceDocument
    RefreshDocumentLists
End Sub

Private Sub btnPickTarget_Click()
    SelectTargetDocument
    RefreshDocumentLists
End Sub

Private Sub btnCheckParas_Click()
    CheckParagraphCount
End Sub

'--- Tab Switching ---
Private Sub tabMain_Change()
    ShowTab tabMain.Value
End Sub

Private Sub ShowTab(idx As Integer)
    fraLayout.Visible = (idx = 0)
    fraFormat.Visible = (idx = 1)
    fraADP.Visible = (idx = 2)
End Sub

'--- Layout Tab ---
Private Sub btnApplyLayout_Click()
    CopyPageLayout chkPageSize.Value, chkHeader.Value, chkFooter.Value, _
                   chkCopyHFContent.Value, chkDeleteEmptyParas.Value
End Sub

Private Sub chkCopyHFContent_Click()
    chkDeleteEmptyParas.Visible = chkCopyHFContent.Value
End Sub

Private Sub btnCopyStyles_Click()
    RunCopyStyles
End Sub

'--- Format Tab ---
Private Sub btnFormatHelper_Click()
    RunFormatHelper
End Sub

Private Sub btnCleanMEPS_Click()
    CleanMEPS chkMEPSWTS.Value, chkMEPSUniversal.Value
End Sub

'--- ADP Tab ---
Private Sub btnSelectExcel_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select Excel Dictionary"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx;*.xls"
    If fd.Show = -1 Then
        lblExcelFile.Caption = fd.SelectedItems(1)
        SaveSetting_ "LastExcelFile", fd.SelectedItems(1)
    End If
End Sub

Private Sub btnRunReplace_Click()
    ReplaceKeyPhrases chkSkipComments.Value, chkIncludeHF.Value
End Sub

Private Sub btnTransferDP_Click()
    TransferDPComments txtFindId.Text, txtReplaceWith.Text, chkExcludeIt.Value
End Sub

Private Sub btnDetectMissing_Click()
    RunDetectMissingSections
End Sub

'--- Batch Mode ---
Private Sub chkBatchMode_Click()
    btnSelectFiles.Visible = chkBatchMode.Value
    btnSelectFolder.Visible = chkBatchMode.Value
End Sub

Private Sub btnSelectFiles_Click()
    RunBatchProcess
End Sub

Private Sub btnSelectFolder_Click()
    RunBatchProcess
End Sub
```
