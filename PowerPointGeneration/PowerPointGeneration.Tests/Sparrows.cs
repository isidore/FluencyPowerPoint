//using System.Data;

using System;
using System.IO;
using System.Linq;
using System.Net;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.Utilities;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointGeneration.Tests
{
	public class Sparrows
	{
		public static void Create()
		{
			DownloadAllFiles();
			Application pptApplication = new Application();
			// Create the Presentation File
			Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
			AddSparrows(pptPresentation);
			pptPresentation.SaveAs(@"c:\temp\Sparrows.pptx", PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
			pptPresentation.Close();
//			pptApplication.Quit();
		}

		private static void DownloadAllFiles()
		{
			int counter = 0;
			foreach (var sparrow in SparrowData.Get())
			{
				counter++;
				string localFilename = @"c:\temp\birds\sparrow{0}.jpg".FormatWith(counter);
				if (!File.Exists(localFilename))
				{
					using (WebClient client = new WebClient())
					{
						client.DownloadFile(sparrow.Item2, localFilename);
					}
				}
			}
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
				foreach (var sparrow in SparrowData.Get())
				{
					counter++;
					// Question
					totalTime = AddPicturePage(slides, page, customLayout, counter, totalTime);
					page++;
					// Answer
					totalTime = AddAnswerPage(slides, page, textLayout, sparrow, counter, totalTime);
					page++;
				}
				Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime %60));
			}
		}

		private static float AddPicturePage(Slides slides, int page, CustomLayout customLayout, int counter, float totalTime)
		{
			var slide = slides.AddSlide(page, customLayout);
			Shape shape = slide.Shapes[2];
			string localFilename = @"c:\temp\birds\sparrow{0}.jpg".FormatWith(counter);
			slide.Shapes.AddPicture(localFilename, MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top,
				shape.Width, shape.Height);
			float time = GetTimingsForImage(counter);
			totalTime += time;
			slide.SlideShowTransition.AdvanceTime = time;
			slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
			return totalTime;
		}

		private static float AddAnswerPage(Slides slides, int page, CustomLayout textLayout, Tuple<string, string> sparrow, int counter,
			float totalTime)
		{
			Slide slide;
			float time;
			slide = slides.AddSlide(page, textLayout);
			var title = slide.Shapes[1].TextFrame.TextRange;
			title.Text = sparrow.Item1.Split('\n').First();
			title.Font.Name = "Arial";
			title.Font.Size = 80;
			var subtitle = slide.Shapes[2].TextFrame.TextRange;
			subtitle.Text = sparrow.Item1.Split('\n').Last();
			subtitle.Font.Name = "Arial";
			subtitle.Font.Size = 30;
			time = GetTimingsForAnswer(counter);
			totalTime += time;
			slide.SlideShowTransition.AdvanceTime = time;
			slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
			return totalTime;
		}

		private static float GetTimingsForImage(int counter)
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

		private static float GetTimingsForAnswer(int counter)
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