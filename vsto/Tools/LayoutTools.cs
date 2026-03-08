using System;
using Word = Microsoft.Office.Interop.Word;

namespace SDTD.Tools
{
    public static class LayoutTools
    {
        /// <summary>
        /// Copies page layout (size, headers, footers) from source to target.
        /// </summary>
        public static void ApplyLayout(Word.Document src, Word.Document tgt, SdtdRibbon opts)
        {
            var srcSetup = src.PageSetup;
            var tgtSetup = tgt.PageSetup;

            if (opts.OptPageSize)
            {
                tgtSetup.PaperSize = srcSetup.PaperSize;
                tgtSetup.PageWidth = srcSetup.PageWidth;
                tgtSetup.PageHeight = srcSetup.PageHeight;
                tgtSetup.Orientation = srcSetup.Orientation;
                tgtSetup.TopMargin = srcSetup.TopMargin;
                tgtSetup.BottomMargin = srcSetup.BottomMargin;
                tgtSetup.LeftMargin = srcSetup.LeftMargin;
                tgtSetup.RightMargin = srcSetup.RightMargin;
            }

            // Process each section
            int sectionCount = Math.Min(src.Sections.Count, tgt.Sections.Count);
            for (int i = 1; i <= sectionCount; i++)
            {
                var srcSection = src.Sections[i];
                var tgtSection = tgt.Sections[i];

                if (opts.OptHeader)
                    CopyHeaderFooter(srcSection, tgtSection, isHeader: true,
                        copyContent: opts.OptHFContent, deleteEmpty: opts.OptDelEmptyParas);

                if (opts.OptFooter)
                    CopyHeaderFooter(srcSection, tgtSection, isHeader: false,
                        copyContent: opts.OptHFContent, deleteEmpty: opts.OptDelEmptyParas);
            }
        }

        private static void CopyHeaderFooter(Word.Section src, Word.Section tgt,
            bool isHeader, bool copyContent, bool deleteEmpty)
        {
            var hfTypes = new[]
            {
                Word.WdHeaderFooterIndex.wdHeaderFooterPrimary,
                Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages
            };

            foreach (var hfType in hfTypes)
            {
                try
                {
                    Word.HeaderFooter srcHF, tgtHF;
                    if (isHeader)
                    {
                        srcHF = src.Headers[hfType];
                        tgtHF = tgt.Headers[hfType];
                    }
                    else
                    {
                        srcHF = src.Footers[hfType];
                        tgtHF = tgt.Footers[hfType];
                    }

                    tgtHF.LinkToPrevious = srcHF.LinkToPrevious;

                    if (copyContent)
                    {
                        // Clear target header/footer
                        tgtHF.Range.Delete();

                        // Copy source content
                        srcHF.Range.Copy();
                        tgtHF.Range.Paste();

                        if (deleteEmpty)
                            DeleteEmptyParagraphs(tgtHF.Range);
                    }
                }
                catch { /* Some header/footer types may not exist */ }
            }
        }

        private static void DeleteEmptyParagraphs(Word.Range range)
        {
            for (int i = range.Paragraphs.Count; i >= 1; i--)
            {
                var para = range.Paragraphs[i];
                string text = para.Range.Text.Trim();
                // Paragraph mark is \r — empty paragraph is just that
                if (string.IsNullOrEmpty(text) || text == "\r")
                    para.Range.Delete();
            }
        }

        /// <summary>
        /// Copies all paragraph styles from source to target document.
        /// </summary>
        public static void CopyStyles(Word.Document src, Word.Document tgt)
        {
            foreach (Word.Style style in src.Styles)
            {
                try
                {
                    if (style.Type == Word.WdStyleType.wdStyleTypeParagraph ||
                        style.Type == Word.WdStyleType.wdStyleTypeCharacter)
                    {
                        // Try to copy the style by applying source style properties
                        Word.Style tgtStyle;
                        try
                        {
                            tgtStyle = tgt.Styles[style.NameLocal];
                        }
                        catch
                        {
                            // Style doesn't exist in target — add it
                            tgtStyle = tgt.Styles.Add(style.NameLocal, style.Type);
                        }

                        // Copy font properties
                        tgtStyle.Font.Name = style.Font.Name;
                        tgtStyle.Font.Size = style.Font.Size;
                        tgtStyle.Font.Bold = style.Font.Bold;
                        tgtStyle.Font.Italic = style.Font.Italic;
                        tgtStyle.Font.Color = style.Font.Color;

                        // Copy paragraph format for paragraph styles
                        if (style.Type == Word.WdStyleType.wdStyleTypeParagraph)
                        {
                            tgtStyle.ParagraphFormat.Alignment = style.ParagraphFormat.Alignment;
                            tgtStyle.ParagraphFormat.SpaceBefore = style.ParagraphFormat.SpaceBefore;
                            tgtStyle.ParagraphFormat.SpaceAfter = style.ParagraphFormat.SpaceAfter;
                            tgtStyle.ParagraphFormat.LineSpacingRule = style.ParagraphFormat.LineSpacingRule;
                            tgtStyle.ParagraphFormat.LineSpacing = style.ParagraphFormat.LineSpacing;
                            tgtStyle.ParagraphFormat.FirstLineIndent = style.ParagraphFormat.FirstLineIndent;
                            tgtStyle.ParagraphFormat.LeftIndent = style.ParagraphFormat.LeftIndent;
                            tgtStyle.ParagraphFormat.RightIndent = style.ParagraphFormat.RightIndent;
                        }
                    }
                }
                catch { /* Skip built-in styles that can't be modified */ }
            }
        }
    }
}
