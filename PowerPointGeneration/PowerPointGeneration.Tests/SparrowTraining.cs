﻿using System;
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
    public class SparrowTraining
    {
        public static void Create(string houseText, string songText)
        {
            Application pptApplication = new Application();
            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            AddSparrows(pptPresentation, GetTrainingSet(houseText, songText));
            pptPresentation.SaveAs(@"c:\temp\Sparrows.pptx", PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            pptPresentation.Close();
            //			pptApplication.Quit();
        }

        public static void Create()
        {
            Create("House", "Song");
        }


        private static Tuple<string, string>[] GetTrainingSet(string houseText, string songText)
        {
            var house =
                Enumerable.Range(1, 53)
                    .Select(n => Tuple.Create(houseText, @"c:\temp\birds\sparrow_house_{0:00}.jpg".FormatWith(n)))
                    .ToArray();
            var chipping =
                Enumerable.Range(1, 32)
                    .Select(n => Tuple.Create("Chipping", @"c:\temp\birds\sparrow_chipping_{0:00}.jpg".FormatWith(n)))
                    .ToArray();
            var song = Enumerable.Range(1, 53)
                .Select(n => Tuple.Create(songText, @"c:\temp\birds\sparrow_song_{0:00}.jpg".FormatWith(n)))
                .ToArray();

            int amount = 53;
            return CreateShuffledDeck(house.Take(amount), song.Take(amount));
        }

        private static T[] CreateShuffledDeck<T>(IEnumerable<T> deckA, IEnumerable<T> deckB)
        {
            var results = new List<T>();
            var listA = new Queue<T>(deckA);
            var listB = new Queue<T>(deckB);
            results.Add(listA.Dequeue());
            results.Add(listB.Dequeue());
            var random = new Random();
            while (0 < listA.Count + listB.Count)
            {
                var queue = (random.NextDouble() < 0.5) ? listA : listB;
                if (0 < queue.Count)
                {
                    results.Add(queue.Dequeue());
                }
            }
            return results.ToArray();
        }


        private static void AddSparrows(Presentation pptPresentation, Tuple<string, string>[] trainingSet)
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
                foreach (var sparrow in trainingSet)
                {
                    // Question
                    totalTime = AddPicturePage(slides, page, customLayout, sparrow, counter, totalTime);
                    page++;
                    // Answer
                    totalTime = AddAnswerPage(slides, page, textLayout, sparrow, counter, totalTime);
                    page++;
                    counter++;
                }
                Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime%60));
            }
        }

        private static float AddPicturePage(Slides slides, int page, CustomLayout customLayout,
            Tuple<string, string> sparrow, int counter, float totalTime)
        {
            var slide = slides.AddSlide(page, customLayout);
            PlaceImageOnPage(sparrow, slide);
            float time = GetTimingsForImage(counter);
            totalTime += time;
            slide.SlideShowTransition.AdvanceTime = time;
            slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
            return totalTime;
        }
        private static void PlaceImageOnPage(Tuple<string, string> sparrow, Slide slide)
        {
            var slideHeight = slide.Design.SlideMaster.Height;
            var slideWidth = slide.Design.SlideMaster.Width;
            var shape = getShapeSizing(sparrow, slide, slideHeight, slideWidth);

            slide.Shapes.AddPicture(sparrow.Item2, MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top,
                shape.Width, shape.Height);
        }
        private static Shape getShapeSizing(Tuple<string, string> sparrow, Slide slide, float slideHeight, float slideWidth)
        {
            Image image = Image.FromFile(sparrow.Item2);
            Shape shape = slide.Shapes[2];
            var imageWidth = image.Width;
            var imageHeight = image.Height;
            if (imageHeight < imageWidth)
            {
                shape.Height = imageHeight * (slideWidth / (float)imageWidth);
                shape.Width = slideWidth;
                shape.Top = (slideHeight - shape.Height) / 2.0F;
                shape.Left = 0;
            }
            else
            {
                shape.Width = imageWidth * (slideHeight / (float)imageHeight);
                shape.Height = slideHeight;
                shape.Top = 0;
                shape.Left = (slideWidth - shape.Width) / 2.0F;
            }
            //            Logger.Variable("Shape",
            //                "[top={0},left={1},{2},{3}]".FormatWith(shape.Top, shape.Left, shape.Width, shape.Height));
            return shape;
        }
        private static float AddAnswerPage(Slides slides, int page, CustomLayout textLayout,
            Tuple<string, string> sparrow, int counter,
            float totalTime)
        {
            var slide = slides.AddSlide(page, textLayout);
            PlaceImageOnPage(sparrow,slide);
            PlaceTextOnPage(sparrow, slide);

            float time = GetTimingsForAnswer(counter);
            totalTime += time;
            slide.SlideShowTransition.AdvanceTime = time;
            slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
            return totalTime; 
            
        }

        private static void PlaceTextOnPage(Tuple<string, string> sparrow, Slide slide)
        {
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = sparrow.Item1;
            title.Font.Name = "Arial Black";
            title.Font.Size = 80;
            slide.Shapes[1].Top = 0;
            slide.Shapes[1].Left = 0;
            slide.Shapes[1].Width = slide.Design.SlideMaster.Width;
            slide.Shapes[1].Height = slide.Design.SlideMaster.Height;
            var color = 0xFFFFFF;
            title.Font.Color.RGB = color;
            title.Font.Shadow = MsoTriState.msoTrue;
            slide.Shapes[1].ZOrder(MsoZOrderCmd.msoBringToFront);
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