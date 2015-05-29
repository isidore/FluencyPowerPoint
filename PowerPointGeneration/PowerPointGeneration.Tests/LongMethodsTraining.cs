//using System.Data;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.Utilities;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointGeneration.Tests
{
	public class LongMethodsTraining
	{
		public static void Create()
		{
			Application pptApplication = new Application();
			// Create the Presentation File
			Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
			AddCode(pptPresentation);
			pptPresentation.SaveAs(@"c:\temp\LongMethods.pptx", PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
			pptPresentation.Close();
		}

		private static Tuple<string, string>[] GetTrainingSet()
		{
			var dir = new DirectoryInfo(@"C:\code\FluencyPowerPoint\PowerPointGeneration\PowerPointGeneration.Tests\LongMethodSmells");
			var files = dir.GetFiles("*.png");
			var longMethods = files.Select(f => Tuple.Create(f.FullName, "Too Long"));

			dir = new DirectoryInfo(@"C:\code\FluencyPowerPoint\PowerPointGeneration\PowerPointGeneration.Tests\LongMethodSmells\Short Methods");
			files = dir.GetFiles("*.png");
			var shortMethods = files.Select(f => Tuple.Create(f.FullName, "Short Enough"));

			return shortMethods.Concat(longMethods).Shuffle();
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
					totalTime = AddPicturePage(slides, page, customLayout, code, counter, totalTime);
					page++;
					// Answer
					totalTime = AddAnswerPage(slides, page, textLayout, code, counter, totalTime);
					page++;
				}
				Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime%60));
			}
		}

		private static float AddPicturePage(Slides slides, int page, CustomLayout customLayout,
			Tuple<string, string> code, int counter, float totalTime)
		{
			var slide = slides.AddSlide(page, customLayout);
			Shape shape = slide.Shapes[2];
			slide.Shapes.AddPicture(code.Item1, MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top,
				shape.Width, shape.Height);
			slide.Background.Fill.ForeColor.RGB = 0x2D2D2D;
			slide.FollowMasterBackground = MsoTriState.msoFalse;
			float time = GetTimingsForImage(counter);
			totalTime += time;
			slide.SlideShowTransition.AdvanceTime = time;
			slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
			return totalTime;
		}

		private static float AddAnswerPage(Slides slides, int page, CustomLayout textLayout,
			Tuple<string, string> code, int counter,
			float totalTime)
		{
			Slide slide;
			float time;
			slide = slides.AddSlide(page, textLayout);
			slide.Background.Fill.ForeColor.RGB = 0x2D2D2D;
			slide.FollowMasterBackground = MsoTriState.msoFalse;
			var title = slide.Shapes[1].TextFrame.TextRange;
			title.Text = code.Item2;
			title.Font.Name = "Arial";
			title.Font.Size = 80;
			var color = code.Item2.Contains("Long") ? 0x3B3BFF : 0x6AE869;
			title.Font.Color.RGB = color;
			time = GetTimingsForAnswer(counter);
			totalTime += time;
			slide.SlideShowTransition.AdvanceTime = time;
			slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
			return totalTime;
		}

		public static float GetTimingsForImage(int counter)
		{
			if (counter < 5)
			{
				return 4.0f;
			}
			else if (counter < 12)
			{
				return 3.0f;
			}
			else if (counter < 20)
			{
				return 2.0f;
			}
			else if (counter < 30)
			{
				return 1.5f;
			}
			return 1f;
		}

		public static float GetTimingsForAnswer(int counter)
		{
			if (counter < 5)
			{
				return 1.5f;
			}
			else if (counter < 12)
			{
				return 1;
			}
			else if (counter < 20)
			{
				return 0.75f;
			}
			else if (counter < 30)
			{
				return 0.5f;
			}
			return 0.5f;
		}
	}

}