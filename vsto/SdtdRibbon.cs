using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using SDTD.Tools;
using Word = Microsoft.Office.Interop.Word;

namespace SDTD
{
    [ComVisible(true)]
    public class SdtdRibbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;
        private readonly ThisAddIn _addIn;

        // State
        private List<string> _openDocs = new List<string>();
        private int _sourceIndex = -1;
        private int _targetIndex = -1;

        // Layout options
        internal bool OptPageSize = true;
        internal bool OptHeader = true;
        internal bool OptFooter = true;
        internal bool OptHFContent = false;
        internal bool OptDelEmptyParas = false;

        // Format options
        internal bool OptMepsBWTS = true;
        internal bool OptMepsBUniv = true;

        // ADP options
        internal string ExcelFilePath = null;
        internal bool OptSkipComments = false;
        internal bool OptIncludeHF = false;
        internal string FindId = "=E";
        internal string ReplaceWith = "=BL";
        internal bool OptExcludeIt = false;

        // Batch
        internal bool BatchMode = false;
        internal List<string> BatchFiles = new List<string>();

        public SdtdRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("SDTD.SdtdRibbon.xml"))
            using (var reader = new StreamReader(stream))
                return reader.ReadToEnd();
        }

        #endregion

        #region Callbacks

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            RefreshDocList();
        }

        // ── Document dropdowns ──────────────────────────────────────

        private void RefreshDocList()
        {
            _openDocs.Clear();
            foreach (Word.Document doc in _addIn.WordApp.Documents)
                _openDocs.Add(doc.Name);
            _ribbon?.InvalidateControl("ddSource");
            _ribbon?.InvalidateControl("ddTarget");
        }

        public int GetSourceCount(IRibbonControl ctrl) => _openDocs.Count;
        public string GetSourceLabel(IRibbonControl ctrl, int index) => _openDocs[index];
        public void OnSourceSelect(IRibbonControl ctrl, string id, int index) { _sourceIndex = index; }

        public int GetTargetCount(IRibbonControl ctrl) => _openDocs.Count;
        public string GetTargetLabel(IRibbonControl ctrl, int index) => _openDocs[index];
        public void OnTargetSelect(IRibbonControl ctrl, string id, int index) { _targetIndex = index; }

        public void OnRefreshDocs(IRibbonControl ctrl) => RefreshDocList();

        internal Word.Document GetSourceDoc()
        {
            if (_sourceIndex < 0 || _sourceIndex >= _openDocs.Count)
            { MessageBox.Show("Please select a Source document.", "SDTD"); return null; }
            return _addIn.WordApp.Documents[_openDocs[_sourceIndex]];
        }

        internal Word.Document GetTargetDoc()
        {
            if (_targetIndex < 0 || _targetIndex >= _openDocs.Count)
            { MessageBox.Show("Please select a Target document.", "SDTD"); return null; }
            return _addIn.WordApp.Documents[_openDocs[_targetIndex]];
        }

        // ── Paragraph Count ─────────────────────────────────────────

        public void OnCheckParagraphCount(IRibbonControl ctrl)
        {
            var src = GetSourceDoc(); var tgt = GetTargetDoc();
            if (src == null || tgt == null) return;
            int sCount = src.Paragraphs.Count;
            int tCount = tgt.Paragraphs.Count;
            string status = sCount == tCount ? "✓ Match" : "⚠ Mismatch";
            MessageBox.Show($"Source: {sCount} paragraphs\nTarget: {tCount} paragraphs\n\n{status}",
                "Paragraph Count", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ── Layout checkboxes ───────────────────────────────────────

        public bool GetChkPageSize(IRibbonControl c) => OptPageSize;
        public void OnChkPageSize(IRibbonControl c, bool v) { OptPageSize = v; }
        public bool GetChkHeader(IRibbonControl c) => OptHeader;
        public void OnChkHeader(IRibbonControl c, bool v) { OptHeader = v; }
        public bool GetChkFooter(IRibbonControl c) => OptFooter;
        public void OnChkFooter(IRibbonControl c, bool v) { OptFooter = v; }
        public bool GetChkHFContent(IRibbonControl c) => OptHFContent;
        public void OnChkHFContent(IRibbonControl c, bool v) { OptHFContent = v; }
        public bool GetChkDelEmpty(IRibbonControl c) => OptDelEmptyParas;
        public void OnChkDelEmpty(IRibbonControl c, bool v) { OptDelEmptyParas = v; }

        // ── Layout actions ──────────────────────────────────────────

        public void OnApplyLayout(IRibbonControl ctrl)
        {
            var src = GetSourceDoc(); var tgt = GetTargetDoc();
            if (src == null || tgt == null) return;
            RunWithStatus("Apply Layout", () => LayoutTools.ApplyLayout(src, tgt, this));
        }

        public void OnCopyStyles(IRibbonControl ctrl)
        {
            var src = GetSourceDoc(); var tgt = GetTargetDoc();
            if (src == null || tgt == null) return;
            RunWithStatus("Copy Styles", () => LayoutTools.CopyStyles(src, tgt));
        }

        // ── Format checkboxes & actions ─────────────────────────────

        public bool GetChkMepsBWTS(IRibbonControl c) => OptMepsBWTS;
        public void OnChkMepsBWTS(IRibbonControl c, bool v) { OptMepsBWTS = v; }
        public bool GetChkMepsBUniv(IRibbonControl c) => OptMepsBUniv;
        public void OnChkMepsBUniv(IRibbonControl c, bool v) { OptMepsBUniv = v; }

        public void OnFormatHelper(IRibbonControl ctrl)
        {
            var src = GetSourceDoc(); var tgt = GetTargetDoc();
            if (src == null || tgt == null) return;
            RunWithStatus("Format Helper", () => FormattingTools.RunFormatHelper(src, tgt));
        }

        public void OnCleanMeps(IRibbonControl ctrl)
        {
            var tgt = GetTargetDoc();
            if (tgt == null) return;
            var fonts = new List<string>();
            if (OptMepsBWTS) fonts.Add("MEPS Bookman WTS");
            if (OptMepsBUniv) fonts.Add("MEPS Bookman Universal");
            if (fonts.Count == 0) { MessageBox.Show("Select at least one MEPS font.", "SDTD"); return; }
            RunWithStatus("Clean MEPS Up", () => FormattingTools.CleanMeps(tgt, fonts));
        }

        // ── ADP checkboxes & actions ────────────────────────────────

        public bool GetChkSkipComments(IRibbonControl c) => OptSkipComments;
        public void OnChkSkipComments(IRibbonControl c, bool v) { OptSkipComments = v; }
        public bool GetChkIncludeHF(IRibbonControl c) => OptIncludeHF;
        public void OnChkIncludeHF(IRibbonControl c, bool v) { OptIncludeHF = v; }
        public string GetFindId(IRibbonControl c) => FindId;
        public void OnFindIdChange(IRibbonControl c, string v) { FindId = v; }
        public string GetReplaceWith(IRibbonControl c) => ReplaceWith;
        public void OnReplaceWithChange(IRibbonControl c, string v) { ReplaceWith = v; }
        public bool GetChkExcludeIt(IRibbonControl c) => OptExcludeIt;
        public void OnChkExcludeIt(IRibbonControl c, bool v) { OptExcludeIt = v; }

        public void OnSelectExcel(IRibbonControl ctrl)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel Files|*.xlsx;*.xls";
                dlg.Title = "Select Dictionary Excel File";
                if (dlg.ShowDialog() == DialogResult.OK)
                    ExcelFilePath = dlg.FileName;
            }
        }

        public void OnReplaceKeyPhrases(IRibbonControl ctrl)
        {
            var tgt = GetTargetDoc();
            if (tgt == null) return;
            if (string.IsNullOrEmpty(ExcelFilePath))
            { MessageBox.Show("Please select an Excel dictionary file first.", "SDTD"); return; }
            RunWithStatus("Replace Key Phrases", () =>
                AdpTools.ReplaceKeyPhrases(tgt, ExcelFilePath, OptSkipComments, OptIncludeHF));
        }

        public void OnTransferComments(IRibbonControl ctrl)
        {
            var src = GetSourceDoc(); var tgt = GetTargetDoc();
            if (src == null || tgt == null) return;
            RunWithStatus("Transfer DP Comments", () =>
                AdpTools.TransferDPComments(src, tgt, FindId, ReplaceWith, OptExcludeIt));
        }

        public void OnDetectMissing(IRibbonControl ctrl)
        {
            var src = GetSourceDoc(); var tgt = GetTargetDoc();
            if (src == null || tgt == null) return;
            RunWithStatus("Detect Missing Sections", () =>
                AdpTools.DetectMissingSections(src, tgt));
        }

        // ── Batch ───────────────────────────────────────────────────

        public bool GetChkBatch(IRibbonControl c) => BatchMode;
        public void OnChkBatch(IRibbonControl c, bool v) { BatchMode = v; }

        public void OnBatchSelectFiles(IRibbonControl ctrl)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Word Documents|*.docx;*.doc";
                dlg.Multiselect = true;
                dlg.Title = "Select Target Files for Batch Processing";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    BatchFiles = dlg.FileNames.ToList();
                    MessageBox.Show($"Selected {BatchFiles.Count} file(s).", "Batch Mode");
                }
            }
        }

        public void OnBatchSelectFolder(IRibbonControl ctrl)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select folder with target .docx files";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    BatchFiles = Directory.GetFiles(dlg.SelectedPath, "*.docx").ToList();
                    BatchFiles.AddRange(Directory.GetFiles(dlg.SelectedPath, "*.doc"));
                    MessageBox.Show($"Found {BatchFiles.Count} file(s) in folder.", "Batch Mode");
                }
            }
        }

        // ── Helpers ─────────────────────────────────────────────────

        private void RunWithStatus(string label, Action action)
        {
            _addIn.WordApp.StatusBar = $"SDTD: Running {label}...";
            _addIn.WordApp.ScreenUpdating = false;
            try
            {
                action();
                _addIn.WordApp.StatusBar = $"SDTD: {label} completed.";
                MessageBox.Show($"{label} completed successfully.", "SDTD",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{label} failed:\n{ex.Message}", "SDTD Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                _addIn.WordApp.StatusBar = $"SDTD: {label} failed.";
            }
            finally
            {
                _addIn.WordApp.ScreenUpdating = true;
            }
        }

        #endregion
    }
}
