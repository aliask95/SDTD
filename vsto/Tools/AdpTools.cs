using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace SDTD.Tools
{
    public static class AdpTools
    {
        /// <summary>
        /// Reads an Excel dictionary (Column A → Column B) and replaces all
        /// matching phrases in the Target document.
        /// Requires EPPlus NuGet package.
        /// </summary>
        public static void ReplaceKeyPhrases(Word.Document tgt, string excelPath,
            bool skipComments, bool includeHeaderFooter)
        {
            // Read Excel dictionary
            var replacements = ReadExcelDictionary(excelPath);
            if (replacements.Count == 0)
            {
                MessageBox.Show("No replacement pairs found in the Excel file.", "SDTD");
                return;
            }

            int replaced = 0;

            // Replace in main body
            foreach (var kvp in replacements)
            {
                replaced += FindAndReplace(tgt.Content, kvp.Key, kvp.Value);
            }

            // Replace in headers/footers if requested
            if (includeHeaderFooter)
            {
                foreach (Word.Section section in tgt.Sections)
                {
                    foreach (Word.HeaderFooter hf in section.Headers)
                        foreach (var kvp in replacements)
                            replaced += FindAndReplace(hf.Range, kvp.Key, kvp.Value);

                    foreach (Word.HeaderFooter hf in section.Footers)
                        foreach (var kvp in replacements)
                            replaced += FindAndReplace(hf.Range, kvp.Key, kvp.Value);
                }
            }

            // Optionally handle comments
            if (!skipComments)
            {
                foreach (Word.Comment comment in tgt.Comments)
                {
                    foreach (var kvp in replacements)
                        replaced += FindAndReplace(comment.Range, kvp.Key, kvp.Value);
                }
            }

            MessageBox.Show($"Replace Key Phrases complete.\n{replaced} replacement(s) made.",
                "SDTD", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static Dictionary<string, string> ReadExcelDictionary(string path)
        {
            var dict = new Dictionary<string, string>();

            // Using EPPlus to read Excel
            // Add NuGet: EPPlus
            try
            {
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(path)))
                {
                    var sheet = package.Workbook.Worksheets[0];
                    int row = 1;
                    while (sheet.Cells[row, 1].Value != null)
                    {
                        string find = sheet.Cells[row, 1].Value?.ToString() ?? "";
                        string replace = sheet.Cells[row, 2].Value?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(find))
                            dict[find] = replace;
                        row++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel file:\n{ex.Message}", "SDTD",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dict;
        }

        private static int FindAndReplace(Word.Range range, string find, string replace)
        {
            int count = 0;
            var findObj = range.Find;
            findObj.ClearFormatting();
            findObj.Replacement.ClearFormatting();
            findObj.Text = find;
            findObj.Replacement.Text = replace;
            findObj.Forward = true;
            findObj.Wrap = Word.WdFindWrap.wdFindContinue;
            findObj.Format = false;
            findObj.MatchCase = true;
            findObj.MatchWholeWord = false;
            findObj.MatchWildcards = false;

            while (findObj.Execute(Replace: Word.WdReplace.wdReplaceOne))
                count++;

            return count;
        }

        /// <summary>
        /// Copies comments from author "Digital Publishing" in Source to the corresponding
        /// paragraph in Target. Replaces the Find identifier with the Replace identifier.
        /// </summary>
        public static void TransferDPComments(Word.Document src, Word.Document tgt,
            string findId, string replaceWith, bool excludeIt)
        {
            int transferred = 0;
            int skipped = 0;

            foreach (Word.Comment srcComment in src.Comments)
            {
                if (srcComment.Author != "Digital Publishing") continue;

                string commentText = srcComment.Range.Text;

                // Skip comments containing "it" if option is set
                if (excludeIt && commentText.IndexOf("it", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    skipped++;
                    continue;
                }

                // Replace identifier in comment text
                string newCommentText = commentText.Replace(findId, replaceWith);

                // Find the corresponding paragraph in target
                // Use the paragraph index of the commented range in source
                try
                {
                    int srcParaIndex = GetParagraphIndex(src, srcComment.Scope);
                    if (srcParaIndex > 0 && srcParaIndex <= tgt.Paragraphs.Count)
                    {
                        var tgtPara = tgt.Paragraphs[srcParaIndex];

                        // Also get the commented content text and append to paragraph
                        string scopeText = srcComment.Scope.Text;
                        string modifiedScope = scopeText.Replace(findId, replaceWith);

                        // Add scope text at end of target paragraph (before paragraph mark)
                        Word.Range insertRange = tgtPara.Range.Duplicate;
                        insertRange.Start = insertRange.End - 1; // before ¶ mark
                        insertRange.InsertBefore(modifiedScope);

                        // Add the comment on the inserted text
                        var newComment = tgt.Comments.Add(insertRange, newCommentText);
                        newComment.Author = "Digital Publishing";
                        transferred++;
                    }
                }
                catch { /* paragraph mapping failed, skip */ }
            }

            MessageBox.Show(
                $"Transfer DP Comments complete.\n{transferred} transferred, {skipped} skipped.",
                "SDTD", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static int GetParagraphIndex(Word.Document doc, Word.Range range)
        {
            // Find which paragraph contains this range
            for (int i = 1; i <= doc.Paragraphs.Count; i++)
            {
                var para = doc.Paragraphs[i];
                if (range.Start >= para.Range.Start && range.Start < para.Range.End)
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// Compares headings between Source and Target. Reports headings
        /// present in Source but missing from Target.
        /// </summary>
        public static void DetectMissingSections(Word.Document src, Word.Document tgt)
        {
            var srcHeadings = GetHeadings(src);
            var tgtHeadings = new HashSet<string>(GetHeadings(tgt));

            var missing = srcHeadings.Where(h => !tgtHeadings.Contains(h)).ToList();

            if (missing.Count == 0)
            {
                MessageBox.Show("All source headings are present in the target document.",
                    "Detect Missing Sections", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                string list = string.Join("\n• ", missing);
                MessageBox.Show(
                    $"Found {missing.Count} heading(s) in Source missing from Target:\n\n• {list}",
                    "Detect Missing Sections", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static List<string> GetHeadings(Word.Document doc)
        {
            var headings = new List<string>();
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                try
                {
                    var styleName = ((Word.Style)para.get_Style()).NameLocal;
                    if (styleName.StartsWith("Heading") || styleName.StartsWith("heading"))
                    {
                        string text = para.Range.Text.Trim();
                        if (!string.IsNullOrEmpty(text))
                            headings.Add(text);
                    }
                }
                catch { }
            }
            return headings;
        }
    }
}
