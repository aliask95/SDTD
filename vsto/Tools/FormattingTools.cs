using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace SDTD.Tools
{
    public static class FormattingTools
    {
        /// <summary>
        /// Detects formatted text (Italic, Bold, Bold+Italic) in the Source
        /// and adds review comments in the Target at the corresponding paragraph.
        /// Comment author: "Format Assist"
        /// </summary>
        public static void RunFormatHelper(Word.Document src, Word.Document tgt)
        {
            int srcCount = src.Paragraphs.Count;
            int tgtCount = tgt.Paragraphs.Count;
            int limit = Math.Min(srcCount, tgtCount);
            int commentsAdded = 0;

            for (int i = 1; i <= limit; i++)
            {
                var srcPara = src.Paragraphs[i];
                var formats = DetectFormats(srcPara.Range);

                if (formats.Count > 0)
                {
                    var tgtPara = tgt.Paragraphs[i];
                    string commentText = "Formatting detected: " + string.Join(", ", formats);

                    var comment = tgt.Comments.Add(tgtPara.Range, commentText);
                    comment.Author = "Format Assist";
                    commentsAdded++;
                }
            }

            System.Windows.Forms.MessageBox.Show(
                $"Format Helper complete.\n{commentsAdded} comment(s) added to Target.",
                "Format Helper", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private static List<string> DetectFormats(Word.Range range)
        {
            var formats = new HashSet<string>();

            // Walk through each word/character run to detect formatting
            foreach (Word.Range word in range.Words)
            {
                string text = word.Text.Trim();
                if (string.IsNullOrEmpty(text)) continue;

                bool isBold = IsTrue(word.Bold);
                bool isItalic = IsTrue(word.Italic);

                if (isBold && isItalic)
                    formats.Add($"Bold+Italic(\"{Truncate(text)}\")");
                else if (isBold)
                    formats.Add($"Bold(\"{Truncate(text)}\")");
                else if (isItalic)
                    formats.Add($"Italic(\"{Truncate(text)}\")");
            }

            return new List<string>(formats);
        }

        /// <summary>
        /// Word's Bold/Italic can be 0 (false), -1 (true), or 9999999 (mixed).
        /// We treat -1 as true.
        /// </summary>
        private static bool IsTrue(int value) => value == -1;

        private static string Truncate(string s, int max = 20)
            => s.Length <= max ? s : s.Substring(0, max) + "…";

        /// <summary>
        /// Removes all text runs using the specified MEPS fonts from the Target document.
        /// </summary>
        public static void CleanMeps(Word.Document tgt, List<string> fontNames)
        {
            int removed = 0;

            foreach (string fontName in fontNames)
            {
                var find = tgt.Content.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Font.Name = fontName;
                find.Text = "";
                find.Replacement.Text = "";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindContinue;
                find.Format = true;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;

                // Count matches first for reporting
                Word.Range countRange = tgt.Content.Duplicate;
                while (countRange.Find.Execute(FindText: "", Forward: true,
                    Format: true, Wrap: Word.WdFindWrap.wdFindStop))
                {
                    countRange.Find.Font.Name = fontName;
                    removed++;
                    countRange.Start = countRange.End;
                    if (countRange.Start >= tgt.Content.End) break;
                }

                // Execute replacement (delete all matching text)
                find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }

            System.Windows.Forms.MessageBox.Show(
                $"Clean MEPS Up complete.\nRemoved text in {fontNames.Count} font(s).",
                "Clean MEPS Up", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
}
