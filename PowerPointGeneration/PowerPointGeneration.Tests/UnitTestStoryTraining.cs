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
	public class UnitTestStoryTraining
	{
		public static void Create()
		{
			Application pptApplication = new Application();
			// Create the Presentation File
			Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
			AddCode(pptPresentation);
			pptPresentation.SaveAs(@"c:\temp\UnitTestStories.pptx", PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
			pptPresentation.Close();
		}

		private static Tuple<string, string>[] GetTrainingSet()
		{
			var dir = new DirectoryInfo(@"C:\code\FluencyPowerPoint\PowerPointGeneration\PowerPointGeneration.Tests\UnitTestStories\bad");
			var files = dir.GetFiles("*.png");
			var longMethods = files.Select(f => Tuple.Create(f.FullName, "Not a Story"));

			dir = new DirectoryInfo(@"C:\code\FluencyPowerPoint\PowerPointGeneration\PowerPointGeneration.Tests\UnitTestStories\good");
			files = dir.GetFiles("*.png");
			var shortMethods = files.Select(f => Tuple.Create(f.FullName, "Good Story"));

			return shortMethods.Concat(longMethods).Shuffle().Log("unit test examples", t =>""+ t.Count());
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
			return CodeParagraphsTrainings.AddImage(slides, page, customLayout, totalTime, GetTimingsForImage(counter), code.Item1);
		
		}

		private static float AddAnswerPage(Slides slides, int page, CustomLayout textLayout,
			Tuple<string, string> code, int counter,
			float totalTime)
		{
			Slide slide;
			float time;
			slide = slides.AddSlide(page, textLayout);
			slide.Background.Fill.ForeColor.RGB = 0xFFFFFF;
			slide.FollowMasterBackground = MsoTriState.msoFalse;
			var title = slide.Shapes[1].TextFrame.TextRange;
			title.Text = code.Item2;
			title.Font.Name = "Arial";
			title.Font.Size = 80;
			var color = code.Item2.Contains("Good") ? 0x6AE869 : 0x3B3BFF ;
			title.Font.Color.RGB = color;
			time = GetTimingsForAnswer(counter);
			totalTime += time;
			slide.SlideShowTransition.AdvanceTime = time;
			slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
			return totalTime;
		}

		public static float GetTimingsForImage(int counter)
		{
			var timings = new Dictionary<int, float>{{5, 4.0f}, {12, 3.5f},{25, 3.0f},{Int32.MaxValue, 2.5f}}   ;
		  foreach(var t in timings.OrderBy(t => t.Key))
			{
				if (counter <= t.Key){
				return t.Value;
				}
			}
				throw new Exception("Not found for " + counter);
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