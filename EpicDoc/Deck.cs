namespace EpicDoc
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Linq;

    using ColoredConsole;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

    internal static class Deck
    {
        // Credit: https://www.free-power-point-templates.com/articles/how-to-create-a-powerpoint-presentation-using-c-and-embed-a-picture-to-the-slide/
        internal static void Generate(IEnumerable<EFU> efus)
        {
            ColorConsole.WriteLine($"Generating Deck from Work-items...".Cyan());
            var saveTo = Path.Combine(Environment.CurrentDirectory, "FuncSpec (UserStories).pptx");

            var application = new Application();
            var presentation = application.Presentations.Add(MsoTriState.msoFalse);
            var layout = presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTwoColumnText];
            var slides = presentation.Slides;

            foreach (var item in efus?.Where(x => !x.Workitemtype.Equals("Epic", StringComparison.OrdinalIgnoreCase) && !x.Workitemtype.Equals("Feature", StringComparison.OrdinalIgnoreCase)).Select((efu, i) => new { i = i + 1, efu }))
            {
                var efu = item.efu;
                var slide = slides.AddSlide(item.i, layout);
                var shape = slide.Shapes[1];
                shape.Top = shape.Top - 35;
                shape.Width = (shape.Width / 2) - 5;
                shape.Copy();
                
                var s = slide.Shapes.PasteSpecial(PpPasteDataType.ppPasteShape);
                s.Top = shape.Top;
                s.Left = shape.Left + shape.Width + 10;

                AddText(slide, (efu.Parent.HasValue ? $"[{efu.Parent.Value}] " : string.Empty) + efu.Id + ": " + efu.Title?.Trim(), 64, 42, 832, 100, true);
                AddText(shape, (efu.Description?.TrimEx() ?? "Description?").StripHtml());

                shape = slide.Shapes[3];
                AddText(shape, (efu.AcceptanceCriteria?.TrimEx() ?? "Acceptance Criteria?").StripHtml());

                shape = slide.Shapes[2];
                AddText(shape, "Tags: " + efu.Tags);

                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = efu.Url;
                s.NAR();
                shape.NAR();
                slide.NAR();
            }

            // application.CommandBars.ExecuteMso("SlideZoomInsert");
            
            slides?.NAR();
            presentation.SaveAs(saveTo, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            presentation.Close();
            presentation.NAR();
            application.Quit();
            application.NAR();
            ColorConsole.WriteLine($"Done creating {saveTo}");
            Process.Start("cmd", $"/c \"{saveTo}\"");
        }

        private static void AddText(Slide slide, string text, int left, int top, int width, int height, bool bold = false)
        {
            var shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            AddText(shape, text, bold);
        }

        private static void AddText(Shape shape, string text, bool title = false)
        {
            var tf = shape.TextFrame;
            var tr = tf.TextRange;
            tr.Text = text;
            // tr.Copy();
            // Clipboard.SetTextAsync(text).GetAwaiter().GetResult();
            // ((Application)shape.Application).ActiveWindow.View.PasteSpecial(PpPasteDataType.ppPasteHTML);
            tr.Font.Name = "Segoe UI";
            tr.Font.Size = 12;
            // tr.Font.Bold = title ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            shape.Line.Style = MsoLineStyle.msoLineSingle;
            shape.Line.ForeColor.RGB = Color.DarkGray.ToArgb();
            tf.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            tf.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorNone;
            tf.WordWrap = MsoTriState.msoCTrue;
            // shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            if (title)
            {
                tr.Font.Color.RGB = Color.White.ToArgb();
                shape.Fill.ForeColor.RGB = Color.DimGray.ToArgb();
            }

            tr.NAR();
        }
    }
}