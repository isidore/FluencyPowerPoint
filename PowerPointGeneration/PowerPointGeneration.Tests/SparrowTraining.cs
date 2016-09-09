using System;
using System.Collections.Generic;
using System.Linq;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.Utilities;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointGeneration.Tests
{
    public class SparrowTraining
    {
        public static void Create()
        {
            Application pptApplication = new Application();
            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            AddSparrows(pptPresentation);
            pptPresentation.SaveAs(@"c:\temp\Sparrows.pptx", PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            pptPresentation.Close();
//			pptApplication.Quit();
        }


        private static Tuple<string, string>[] GetTrainingSet()
        {
            var house =
                Enumerable.Range(1, 57)
                    .Select(n => Tuple.Create("House", @"c:\temp\birds\sparrow_house_{0:00}.jpg".FormatWith(n)))
                    .ToArray();
            var chipping =
                Enumerable.Range(1, 32)
                    .Select(n => Tuple.Create("Chipping", @"c:\temp\birds\sparrow_chipping_{0:00}.jpg".FormatWith(n)))
                    .ToArray();
            var song = Enumerable.Range(1, 48)
                .Select(n => Tuple.Create("Song", @"c:\temp\birds\sparrow_song_{0:00}.jpg".FormatWith(n)))
                .ToArray();

            int amount = 48;
            return house.Take(amount).Concat(song.Take(amount)).Shuffle();
        }


        private static void AddSparrows(Presentation pptPresentation)
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
                foreach (var sparrow in GetTrainingSet())
                {
                    counter++;
                    // Question
                    totalTime = AddPicturePage(slides, page, customLayout, sparrow, counter, totalTime);
                    page++;
                    // Answer
                    totalTime = AddAnswerPage(slides, page, textLayout, sparrow, counter, totalTime);
                    page++;
                }
                Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime%60));
            }
        }

        private static float AddPicturePage(Slides slides, int page, CustomLayout customLayout,
            Tuple<string, string> sparrow, int counter, float totalTime)
        {
            var slide = slides.AddSlide(page, customLayout);
            Shape shape = slide.Shapes[2];
            slide.Shapes.AddPicture(sparrow.Item2, MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top,
                shape.Width, shape.Height);
            float time = GetTimingsForImage(counter);
            totalTime += time;
            slide.SlideShowTransition.AdvanceTime = time;
            slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
            return totalTime;
        }

        private static float AddAnswerPage(Slides slides, int page, CustomLayout textLayout,
            Tuple<string, string> sparrow, int counter,
            float totalTime)
        {
            Slide slide;
            float time;
            slide = slides.AddSlide(page, textLayout);
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = sparrow.Item1.Split(' ').First();
            title.Font.Name = "Arial";
            title.Font.Size = 80;
            var subtitle = slide.Shapes[2].TextFrame.TextRange;
            subtitle.Text = sparrow.Item1.Split(' ').Last();
            subtitle.Font.Name = "Arial";
            subtitle.Font.Size = 30;
            time = GetTimingsForAnswer(counter);
            totalTime += time;
            slide.SlideShowTransition.AdvanceTime = time;
            slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
            return totalTime;
        }


        public static float GetTimingsForImage(int counter)
        {
            return new Timings { { 2, 100 }, { 5, 4.0f }, { 12, 3.0f }, { 20, 2.0f }, {30,1.5f} ,{ Int32.MaxValue, 1.0f } }.Get(counter);
        }

        public static float GetTimingsForAnswer(int counter)
        {
            return new Timings {{2, 100}, {5, 1.5f}, {12, 1.0f}, {20, 0.75f}, {Int32.MaxValue, 0.5f}}.Get(counter);
        }
    }

    public static class ArrayUtils
        {
            public static T[] Shuffle<T>(this IEnumerable<T> array)
            {
                Random rnd = new Random();
                return array.OrderBy(x => rnd.Next()).ToArray();
            }
        }
    }