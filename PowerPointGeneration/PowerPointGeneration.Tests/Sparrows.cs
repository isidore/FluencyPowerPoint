//using System.Data;

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

				Slides slides;

				slides = pptPresentation.Slides;
				int counter = 0;
				int page = 1;
				foreach (var sparrow in SparrowData.Get())
				{
					counter++;

					// Question
					var slide = slides.AddSlide(page++, customLayout);
					Shape shape = slide.Shapes[2];
					string localFilename = @"c:\temp\birds\sparrow{0}.jpg".FormatWith(counter);
					slide.Shapes.AddPicture(localFilename, MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top,
						shape.Width, shape.Height);
					float time = GetTimingsForImage(counter);
					totalTime += time;
					slide.SlideShowTransition.AdvanceTime = time;
					slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;

					// Answer
					slide = slides.AddSlide(page++, textLayout);
					var objText = slide.Shapes[1].TextFrame.TextRange;
					objText.Text = sparrow.Item1;
					objText.Font.Name = "Arial";
					objText.Font.Size = 80;
					time = GetTimingsForAnswer(counter);
					totalTime += time;
					slide.SlideShowTransition.AdvanceTime =time;
					slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
				}
				Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime %60));
			}
		}

		private static float GetTimingsForImage(int counter)
		{
			if (counter < 20)
			{
				return 4.0f;
			}
			else if (counter < 30)
			{
				return 3.0f;
			}
			else if (counter < 40)
			{
				return 2.0f;
			}
			else if (counter < 60)
			{
				return 1.5f;
			}
			return 1f;
		}

		private static float GetTimingsForAnswer(int counter)
		{
			return counter < 20 ? 1.5f : 0.75f;
		}
	}
}