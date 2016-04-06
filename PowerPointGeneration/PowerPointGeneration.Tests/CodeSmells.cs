//using System.Data;

using System;
using System.Drawing;
using System.Linq;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.Utilities;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointGeneration.Tests
{
    public class Smell
    {
        public Details Details { get; set; }
        public bool Good { get; set; }

        public string fileName;
        private const string BASE = @"C:\code\FluencyPowerPoint\PowerPointGeneration\PowerPointGeneration.Tests\";

        public Smell(Details details, int number, bool good)
        {
            Details = details;
            Good = good;
            fileName = "CodeSmells-{0}\\{1} {2:00}{3}".FormatWith(details.Name,
                good ? details.GoodName : details.BadName, number, details.FileEndingWithDot);
        }


        internal string GetImage()
        {
            return "{0}{1}".FormatWith(BASE, fileName);
        }
    }

    public class Details
    {
        public int GoodCount { get; set; }
        public int BadCount { get; set; }
        public string Name { get; set; }
        public string GoodName { get; set; }
        public string BadName { get; set; }
        public int BackgroundColor { get; set; }
        public float FontSize { get; set; }
        public string GoodNameText { get; set; }
        public string BadNameText { get; set; }
        public string FileEndingWithDot { get; set; }
        public Timings Timings { get; set; }

        public Details()
        {
            BackgroundColor = 0xFFFFFF;
            FontSize = 120;
            FileEndingWithDot = ".png";
            Timings = new Timings {{2, 100}, {5, 4}, {20, 2.5F}, {Int32.MaxValue, 1.5F}};
        }

        public String GetTextForGood()
        {
            return GoodNameText ?? GoodName;
        }

        public String GetTextForBad()
        {
            return BadNameText ?? BadName;
        }
    }

    public class CodeSmells
    {
        public static void Create(Details details)
        {
            Application pptApplication = new Application();
            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            AddTrainingSet(pptPresentation, details);
            pptPresentation.SaveAs(@"c:\\temp\\{0}.pptx".FormatWith(details.Name), PpSaveAsFileType.ppSaveAsDefault,
                MsoTriState.msoTrue);
            pptPresentation.Close();
        }

        private static Smell[] GetTrainingSet(Details details)
        {
            var files = GetFiles(details);
            return files.Take(2)
                .Concat(files.Skip(2).ToArray().Shuffle())
                .ToArray()
                .Log("# of Examples", v => "" + v.Count());
        }

        private static Smell[] GetFiles(Details details)
        {
            var good = Enumerable.Range(1, details.GoodCount).Select(n => new Smell(details, n, true));
            var bad = Enumerable.Range(1, details.BadCount).Select(n => new Smell(details, n, false));

            return new[] {good.First(), bad.First()}.Concat(good.Skip(1)).Concat(bad.Skip(1)).ToArray();
        }

        private static void AddTrainingSet(Presentation pptPresentation, Details details)
        {
            float totalTime = 0;
            using (Logger.MarkEntryPoints())
            {
                pptPresentation.SlideShowSettings.AdvanceMode = PpSlideShowAdvanceMode.ppSlideShowUseSlideTimings;
                CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];
                CustomLayout textLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];

                Slides slides = pptPresentation.Slides;
                int counter = 0;
                int page = 1;
                foreach (var code in GetTrainingSet(details))
                {
                     // Question
                    totalTime += AddPicturePage(slides, page, customLayout, code, counter);
                    page += 1;


                    // Answer
                    totalTime += AddAnswerPage(slides, page, textLayout, code, counter);
                    page += 1;
                    counter++;
                }
                Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime%60));
            }
        }

        private static float AddAnswerPage(Slides slides, int page, CustomLayout customLayout,
            Smell smell, int counter)
        {
            return AddAnswerImage(slides, page, customLayout, GetTimingsForAnswer(counter), smell);
        }

        private static float AddPicturePage(Slides slides, int page, CustomLayout customLayout,
            Smell smell, int counter)
        {
            float time = GetTimingsForImage(smell.Details, counter);

            return AddImage(slides, page, customLayout, time, smell);
        }

        private static float AddAnswerImage(Slides slides, int page, CustomLayout customLayout, 
            float time, Smell smell)
        {
            var slide = slides.AddSlide(page, customLayout);
            PlaceImageOnPage(smell, slide);
            slide.Background.Fill.ForeColor.RGB = smell.Details.BackgroundColor;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = smell.Good ? smell.Details.GetTextForGood() : smell.Details.GetTextForBad();
            title.Font.Name = "Arial Black";
            title.Font.Size = smell.Details.FontSize;
            slide.Shapes[1].Top = 0;
            slide.Shapes[1].Left = 0;
            slide.Shapes[1].Width = slide.Design.SlideMaster.Width;
            var color = smell.Good ? 0x347400 : 0x3B3BFF;
            title.Font.Color.RGB = color;
            slide.Shapes[1].ZOrder(MsoZOrderCmd.msoBringToFront);
        
            slide.SlideShowTransition.AdvanceTime = time;
            slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
            slide.NotesPage.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 0, 0);
            var t = slide.NotesPage.Shapes[2];
            t.TextFrame.TextRange.Text = smell.fileName;

            return time;
        }

        private static void PlaceImageOnPage(Smell smell, Slide slide)
        {
            var slideHeight = slide.Design.SlideMaster.Height;
            var slideWidth = slide.Design.SlideMaster.Width;
            var shape = getShapeSizing(smell, slide, slideHeight, slideWidth);

            slide.Shapes.AddPicture(smell.GetImage(), MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top,
                shape.Width, shape.Height);
        }

        private static Shape getShapeSizing(Smell smell, Slide slide, float slideHeight, float slideWidth)
        {
            Image image = Image.FromFile(smell.GetImage());
            Shape shape = slide.Shapes[2];
            var imageWidth = image.Width;
            var imageHeight = image.Height;
            if (imageHeight < imageWidth)
            {
                shape.Height = imageHeight*(slideWidth/(float) imageWidth);
                shape.Width = slideWidth;
                shape.Top = (slideHeight - shape.Height)/2.0F;
                shape.Left = 0;
            }
            else
            {
                shape.Width = imageWidth*(slideHeight/(float) imageHeight);
                shape.Height = slideHeight;
                shape.Top = 0;
                shape.Left = (slideWidth - shape.Width)/2.0F;
            }
//            Logger.Variable("Shape",
//                "[top={0},left={1},{2},{3}]".FormatWith(shape.Top, shape.Left, shape.Width, shape.Height));
            return shape;
        }

        public static float AddImage(Slides slides, int page, CustomLayout customLayout, float time, Smell smell)
        {
            var slide = slides.AddSlide(page, customLayout);
            PlaceImageOnPage(smell, slide);
            slide.Background.Fill.ForeColor.RGB = smell.Details.BackgroundColor;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.SlideShowTransition.AdvanceTime = time;
            slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
            return time;
        }


        public static float GetTimingsForImage(Details details, int counter)
        {
            return details.Timings.Get(counter);
        }

        public static float GetTimingsForAnswer(int counter)
        {
            return new Timings { { 2, 100 },{8,1},{ Int32.MaxValue, 0.5f } }.Get(counter); 
        }
    }
}