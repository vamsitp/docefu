namespace EpicDoc
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;

    internal static class Deck
    {
        // Credit: https://www.free-power-point-templates.com/articles/how-to-create-a-powerpoint-presentation-using-c-and-embed-a-picture-to-the-slide/
        internal static void Generate(IEnumerable<EFU> efus)
        {
            var saveTo = "FuncSpec (UserStories).pptx".GetFullPath();

            var application = new Application();
            var presentation = application.Presentations.Add(MsoTriState.msoTrue);
            var layout = presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTwoColumnText];
            Slides slides = presentation.Slides;

            foreach (var item in efus.Where(x => !x.Workitemtype.Equals("Epic", StringComparison.OrdinalIgnoreCase) && !x.Workitemtype.Equals("Feature", StringComparison.OrdinalIgnoreCase)).Select((efu, i) => new { i = i + 1, efu }))
            {
                var efu = item.efu;
                var slide = slides.AddSlide(item.i, layout);

                AddText(slide, efu.Id + ": " + efu.Title?.Trim(), 64, 42, 836, 100, true);

                var shape = slide.Shapes[1];
                AddText(shape, efu.Description?.TrimEx() ?? "Description?");

                shape = slide.Shapes[2];
                AddText(shape, efu.AcceptanceCriteria?.TrimEx() ?? "Acceptance Criteria?");

                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Tags: " + efu.Tags;
                slide.NAR();
            }


            slides?.NAR();
            presentation.SaveAs(saveTo, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            // pptPresentation.Close();
            // pptApplication.Quit();
            // pptApplication.NAR();
        }

        private static void AddText(Slide slide, string text, int left, int top, int width, int height, bool bold = false)
        {
            var shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            AddText(shape, text, bold);
        }

        private static void AddText(Microsoft.Office.Interop.PowerPoint.Shape shape, string text, bool bold = false)
        {
            var tr = shape.TextFrame.TextRange;
            tr.Text = text;
            tr.Font.Name = "Segoe UI";
            tr.Font.Size = 12;
            tr.Font.Bold = bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            shape.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorNone;
            shape.TextFrame.WordWrap = MsoTriState.msoCTrue;
            // shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            tr.NAR();
        }
    }
}