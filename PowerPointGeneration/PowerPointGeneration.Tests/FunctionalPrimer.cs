using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.Utilities;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointGeneration.Tests
{
    public class FunctionalPrimer
    {
        public static void Create()
        {
            FSharpDetails details = new FSharpDetails();
            Application pptApplication = new Application();
            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            AddTrainingSet(pptPresentation, new FSharpDetails());
            pptPresentation.SaveAs(@"c:\\temp\\{0}.pptx".FormatWith(details.Name), PpSaveAsFileType.ppSaveAsDefault,
                MsoTriState.msoTrue);
            pptPresentation.Close();
        }

        private static LanguageFeature[] GetTrainingSet(LanguageGroup group)
        {
            var files = GetFiles(group);
            return files.ToArray().Shuffle();
        }

        private static LanguageFeature[] GetFiles(LanguageGroup details)
        {
            var good = Enumerable.Range(1, details.Count).Select(n => new LanguageFeature(details, n));

            return good.ToArray();
        }

        private static void AddTrainingSet(Presentation pptPresentation, FSharpDetails details)
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
                foreach (var group in details.GetGroups())
                {
                    totalTime += AddTitlePage(slides, page++, customLayout, group);

                    foreach (var code in GetTrainingSet(group))
                    {
//                        // Question
//                        totalTime += AddPicturePage(slides, page, customLayout, code, counter);
//                        page += 1;
//
//
//                        // Answer
//                        totalTime += AddAnswerPage(slides, page, textLayout, code, counter);
//                        page += 1;
//                        counter++;
                    }
                }
                Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime%60));
            }
        }

        private static float AddTitlePage(Slides slides, int page, CustomLayout layout, LanguageGroup group)
        {
            var slide = slides.AddSlide(page, layout);
            slide.Background.Fill.ForeColor.RGB = 0xFFFFFF;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = group.Name;
            title.Font.Name = "Arial Black";
            title.Font.Size = group.TitleSize;
            slide.Shapes[1].Top = 0;
            slide.Shapes[1].Left = 0;
            slide.Shapes[1].Width = slide.Design.SlideMaster.Width;
            slide.Shapes[1].Height = slide.Design.SlideMaster.Height;
            var color = 0x000000;
            title.Font.Color.RGB = color;
            slide.Shapes[1].ZOrder(MsoZOrderCmd.msoBringToFront);

            var time = 1.0f;
            slide.SlideShowTransition.AdvanceTime = time;


            return time;
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
            return new Timings {{2, 100}, {8, 1}, {Int32.MaxValue, 0.5f}}.Get(counter);
        }
    }

    internal class LanguageFeature
    {
        private readonly LanguageGroup details;
        private readonly int number;

        public LanguageFeature(LanguageGroup details, int number)
        {
            this.details = details;
            this.number = number;
        }
    }

    public class FSharpDetails
    {
        public int BackgroundColor = 0xFFFFFF;
        public string Name = "FSharp";

        public IEnumerable<LanguageGroup> GetGroups()
        {
            yield return new LanguageGroup("Types", 8, this);
            yield return new LanguageGroup("Values", 6, this);
            yield return new LanguageGroup("Function", 12, this);
            yield return new LanguageGroup("ForwardPipe", 9, this);
            yield return new LanguageGroup("PatternMatching", 11, this);
            yield return new LanguageGroup("DiscrimatedUnion", 6, this);
        }
    };

    public class LanguageGroup
    {
        public float TitleSize = 48;
        public string Name { get; set; }
        public int Count { get; set; }
        public FSharpDetails Details { get; set; }

        public LanguageGroup(string name, int count, FSharpDetails details)
        {
            this.Name = name;
            this.Count = count;
            Details = details;
        }
    }
}