using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.Utilities;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointGeneration.Tests
{
    public class CodeSmellsBadNames
    {
        public static void Create()
        {
            Application pptApplication = new Application();
            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            AddCode(pptPresentation);
            pptPresentation.SaveAs(@"c:\temp\CodeSmellsBadNames.pptx", PpSaveAsFileType.ppSaveAsDefault,
                MsoTriState.msoTrue);
            pptPresentation.Close();
        }

        private static IEnumerable<Name> GetTrainingSet()
        {
            return GetSamples().Shuffle().Log("Sample Count",x => "" + x.Count());
        }

        private static IEnumerable<Name> GetSamples()
        {
            var text = File.ReadAllText(PathUtilities.GetAdjacentFile("BadNames.trainingset.txt"));
            var parts = text.Replace("\r\n", "\n").Split(new[] {"\n"}, StringSplitOptions.RemoveEmptyEntries);
            Console.WriteLine(parts.ToReadableString());
            for (int i = 0; i < parts.Length - 1; i += 2)
            {
                yield return new Name(parts[i], parts[i + 1]);
            }
        }


        private static void AddCode(Presentation pptPresentation)
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
                foreach (var code in GetTrainingSet())
                {
                    counter++;
                    // Question
                    totalTime = AddSample(slides, page, customLayout, code, counter, totalTime);
                    page += 1;


                    // Answer
                    totalTime = AddAnswerPage(slides, page, customLayout, code, counter, totalTime);
                    page+=1;
                }
                Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime%60));
            }
        }


        private static float AddAnswerPage(Slides slides, int page, CustomLayout customLayout,
            Name sample, int counter, float totalTime)
        {
            float time = GetTimingsForAnswer(counter);
            var slide = AddSampleText(slides, page, customLayout, sample);
            slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, slide.Design.SlideMaster.Width-50, 200);
            
            var answer = slide.Shapes[2].TextFrame.TextRange;
            answer.Text = sample.type + " name";
            answer.Font.Name = "Arial Black";
            answer.Font.Size = 68;
            answer.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            var green = 0x347400;
            var color = sample.type.StartsWith("good") ? green : 0x3B3BFF;
            answer.Font.Color.RGB = color;


            totalTime += time;
            slide.SlideShowTransition.AdvanceTime = time;
            slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;

            return totalTime;
        }

        private static Slide AddSampleText(Slides slides, int page, CustomLayout customLayout, Name sample)
        {
            var slide = slides.AddSlide(page, customLayout);
            slide.Shapes[2].Delete();
            slide.Background.Fill.ForeColor.RGB = 0xFFFFFF;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = sample.NameText;
            title.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
            title.Font.Name = "Verdana";
            title.Font.Size = 28;
            slide.Shapes[1].Top = 0;
            var border = 75;
            slide.Shapes[1].Left = border;
            slide.Shapes[1].Width = slide.Design.SlideMaster.Width - border;
            slide.Shapes[1].Height = slide.Design.SlideMaster.Height;
            var color = 0x000000;
            title.Font.Color.RGB = color;
            slide.Shapes[1].ZOrder(MsoZOrderCmd.msoBringToFront);
            return slide;
        }

        private static float AddSample(Slides slides, int page, CustomLayout customLayout,
            Name sample, int counter, float totalTime)
        {
            float time = GetTimingsForImage(counter);

            var slide = AddSampleText(slides, page, customLayout, sample);
            totalTime += time;
            slide.SlideShowTransition.AdvanceTime = time;
            slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;

            return totalTime;
        }

    

        public static float GetTimingsForImage(int counter)
        {
            return new Timings{{5, 4}, {20, 2}, {Int32.MaxValue, 1}}.Get(counter);
        }


        public static float GetTimingsForAnswer(int counter)
        {
            return 1.5f;
        }
    }

    public class Name
    {
        public readonly string NameText;
        public readonly string type;

        public Name(string name, string type)
        {
            this.NameText = name;
            this.type = type;
        }
    }
}