namespace EpicDoc
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;

    using ColoredConsole;

    using Word = Microsoft.Office.Interop.Word;

    internal class Doc
    {
        private const string HtmlNewLine = "<br/>";
        private static readonly string WordTemplate = ConfigurationManager.AppSettings["WordTemplate"];

        public static void Generate(IEnumerable<EFU> efus)
        {
            string content = string.Join(HtmlNewLine, efus.Select(GetContent));
            ColorConsole.WriteLine($"Generating Document from Work-items...".Cyan());
            var wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone, ScreenUpdating = false };
            object fileName = Path.Combine(Environment.CurrentDirectory, WordTemplate);
            object missing = Type.Missing;
            var wordDoc = wordApp.Documents.Open(
                ref fileName,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);

            var value = "<html><body style=\"font-family:'segoe ui';font-size:14px\">" + content + "</body></html>";
            var bookmark = wordDoc.Bookmarks.get_Item(1);
            ReplaceBookmark(bookmark.Range, value);

            object saveTo = "FuncSpec (UserStories).docx".GetFullPath();
            object format = Word.WdSaveFormat.wdFormatXMLDocument;

            object start = wordDoc.Content.Start;
            object end = wordDoc.Content.End;
            wordDoc.Range(ref start, ref end).Select();
            wordApp.Selection.Range.Font.Name = "Segoe UI";
            wordApp.Selection.Range.Font.Size = 10;

            wordDoc.TrackRevisions = true;
            //// var translate = wordDoc.Research.SetLanguagePair(Word.WdLanguageID.wdSpanishModernSort, Word.WdLanguageID.wdEnglishUS);
            wordDoc.SaveAs(ref saveTo, ref format);
            wordDoc.Close(ref missing, ref missing, ref missing);
            wordDoc.NAR();
            wordApp.Quit(ref missing, ref missing, ref missing);
            wordApp.NAR();
            File.Delete("temp.html".GetFullPath());
            ColorConsole.WriteLine($"Done creating {saveTo}");
            Process.Start("cmd", $"/c \"{saveTo}\"");
        }

        private static string GetContent(EFU efu)
        {
            if (efu == null)
            {
                return string.Empty;
            }

            var desc = string.IsNullOrWhiteSpace(efu.Description) ? string.Empty : efu.Description.TrimEx();
            if (efu.Workitemtype.Equals("Epic", StringComparison.OrdinalIgnoreCase))
            {
                return $"<hr style=\"border:0;height:1px\"/><br/><div style=\"color:#242424\"><b>E-" + efu.Id + ". <u>" + (efu.Title?.ToUpperInvariant() ?? string.Empty) + "</u></b></div>" + desc;
            }

            if (efu.Workitemtype.Equals("Feature", StringComparison.OrdinalIgnoreCase))
            {
                return $"<div style=\"color:#727272\"><b>F-" + efu.Id + ". " + (efu.Title?.ToUpperInvariant() ?? string.Empty) + "</b></div>" + desc;
            }

            var acceptance = string.IsNullOrWhiteSpace(efu.AcceptanceCriteria?.TrimEx()) ? string.Empty : "<b>Acceptance Criteria</b>: " + efu.AcceptanceCriteria.TrimEx();
            return $"<div style=\"color:{Extensions.HeadersColor}\">U-" + efu.Id + ". " + (efu.Title ?? string.Empty) + "</div>" + desc + acceptance;
        }

        public static void ReplaceBookmark(Word.Range rng, string html)
        {
            // var val = string.Format("Version:0.9\nStartHTML:80\nEndHTML:{0,8}\nStartFragment:80\nEndFragment:{0,8}\n", 80 + html.Length) + html + "<";
            // Clipboard.SetData(DataFormats.Html, val);
            //// Clipboard.SetText(val, TextDataFormat.Html);
            // rng.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteHTML);

            rng.Font.Name = "Segoe UI";
            rng.Font.Size = 11;
            var file = "temp.html".GetFullPath();
            File.WriteAllText(file, html);
            rng.InsertFile(file);
        }
    }
}
