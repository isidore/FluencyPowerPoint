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
			foreach (var sparrow in SparrowData.Get())
			{
				string localFilename = GetFileName(sparrow);
				if (!File.Exists(localFilename))
				{
					using (WebClient client = new WebClient())
					{
						client.DownloadFile(sparrow.Item2, localFilename);
					}
				}
			}
		}

		private static Tuple<string, string, int>[] GetTrainingSet()
		{
			var all = SparrowData.Get();
			var house = all.Where(s => s.Item1.StartsWith("House")).ToArray();
			Logger.Variable("House.length", house.Count()); // 63
			var chipping = all.Where(s => s.Item1.StartsWith("Chipping")).ToArray();
			Logger.Variable("Chipping", chipping.Count()); //37
			var song = all.Where(s => s.Item1.StartsWith("Song")).ToArray();
			Logger.Variable("Song", song.Length); // 56
			Random rnd = new Random();
			int amount = 37;
			return house.Take(amount).Concat(chipping.Take(amount)).Concat(song.Take(amount)).OrderBy(x => rnd.Next()).ToArray();
		}

		private static string GetFileName(Tuple<string, string, int> sparrow)
		{
			return @"c:\temp\birds\sparrow{0}.jpg".FormatWith(sparrow.Item3);
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
					totalTime = AddPicturePage(slides, page, customLayout, sparrow,counter, totalTime);
					page++;
					// Answer
					totalTime = AddAnswerPage(slides, page, textLayout, sparrow, counter, totalTime);
					page++;
				}
				Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime %60));
			}
		}

		private static float AddPicturePage(Slides slides, int page, CustomLayout customLayout, Tuple<string, string, int> sparrow, int counter, float totalTime)
		{
			var slide = slides.AddSlide(page, customLayout);
			Shape shape = slide.Shapes[2];
			slide.Shapes.AddPicture(GetFileName(sparrow), MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top,
				shape.Width, shape.Height);
			float time = GetTimingsForImage(counter);
			totalTime += time;
			slide.SlideShowTransition.AdvanceTime = time;
			slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
			return totalTime;
		}

		private static float AddAnswerPage(Slides slides, int page, CustomLayout textLayout, Tuple<string, string, int> sparrow, int counter,
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