using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.Utilities;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointGeneration.Tests
{
	public class SetTraining
	{
		private const int count = 25;
		public static void Create()
		{
			DownloadAllFiles();
			Application pptApplication = new Application();
			// Create the Presentation File
			Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
			AddSets(pptPresentation);
			pptPresentation.SaveAs(@"c:\temp\SetCards.pptx", PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
			pptPresentation.Close();
			//			pptApplication.Quit();
		}


		private static IEnumerable<Set> GetTrainingSet()
		{
			var all = GetAllSets().Shuffle();
			var invalid = all.Where(s => s.GetTypeOfSet() == Set.SetType.Invalid).Take(count*7).ToArray();
			var ones = all.Where(s => s.GetTypeOfSet() == Set.SetType.OneDifference).Take(count*4).ToArray();
			var twos = all.Where(s => s.GetTypeOfSet() == Set.SetType.TwoDifferences).Take(count*4).ToArray();
			var threes = all.Where(s => s.GetTypeOfSet() == Set.SetType.ThreeDifferences).Take(count*3).ToArray();
			var fours = all.Where(s => s.GetTypeOfSet() == Set.SetType.FourDifferences).Take(count*2).ToArray();

			 //1s 2s 1 + 2s 3s + 123 + 4 +1234
			var onesOnly = Part(0, ones).Concat(Part(0,invalid)).Shuffle();
			var twosOnly = Part(0, twos).Concat(Part(1, invalid)).Shuffle();
			var onesAndTwos = Part(1,ones).Concat(Part(1, twos)).Concat(Part(2, invalid)).Shuffle();
			var threesOnly = Part(0, threes).Concat(Part(3, invalid)).Shuffle();
			var onesAndTwosAndThrees = Part(2, ones).Concat(Part(2, twos)).Concat(Part(1, threes)).Concat(Part(4, invalid)).Shuffle();
			var foursOnly = Part(0, fours).Concat(Part(5, invalid)).Shuffle();
			var onesAndTwosAndThreesAndFours = Part(3, ones).Concat(Part(3, twos)).Concat(Part(2, threes)).Concat(Part(6, invalid)).Shuffle();
			
			return onesOnly.Concat(twosOnly).Concat(onesAndTwos).Concat(threesOnly).Concat(onesAndTwosAndThrees).Concat(foursOnly).Concat(onesAndTwosAndThreesAndFours).ToArray();
		}

		private static IEnumerable<Set> Part(int part, IEnumerable<Set> sets)
		{
			return sets.Skip(part*count).Take(count);
		}

		private static IEnumerable<Set> GetAllSets()
		{
			var cards = Enumerable.Range(1, 81).Select(i => new SetCard(i)).ToArray();
			for (int card1 = 0; card1 < 81; card1++)
			{
				for (int card2 = card1 + 1; card2 < 81; card2++)
				{
					for (int card3 = card2 + 1; card3 < 81; card3++)
					{
						yield return new Set(cards[card1], cards[card2], cards[card3]);
					}
				}
			}
		}


		private static void AddSets(Presentation pptPresentation)
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
				foreach (var set in GetTrainingSet())
				{
					counter++;
					// Question
					totalTime = AddPicturePage(slides, page, customLayout, set, counter, totalTime);
					page++;
					// Answer
					totalTime = AddAnswerPage(slides, page, textLayout, set, counter, totalTime);
					page++;
				}
				Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime%60));
			}
		}


		private static void DownloadAllFiles()
		{
			for (int i = 1; i <= 81; i++)
			{
				string localFilename = SetCard.GetImageFileName(i);

				if (!File.Exists(localFilename))
				{
					using (WebClient client = new WebClient())
					{
						string url = "http://www.puzzles.setgame.com/images/setcards/small/{0:00}.gif".FormatWith(i);
						client.DownloadFile(url, localFilename);
					}
				}
			}
		}

		private static float AddPicturePage(Slides slides, int page, CustomLayout customLayout, Set set, int counter,
			float totalTime)
		{
			var slide = slides.AddSlide(page, customLayout);
			for (int i = 1; i < 3; i++)
			{
				slide.Shapes[1].Delete();
			}

			float slideWidth = slide.Application.ActivePresentation.PageSetup.SlideWidth;
			float slideHeight = slide.Application.ActivePresentation.PageSetup.SlideHeight;
			for (int i = 0; i < 3; i++)
			{
				
				float width = 95*1.5f;
				float height = 62*1.5f;
				float left = ((slideWidth / 4) * (i + 1) ) - (width/2);
				slide.Shapes.AddPicture(set.cards[i].GetImageFileName(), MsoTriState.msoFalse, MsoTriState.msoTrue, left,
					slideHeight/2 - (height/2), width,
					height);
				slide.Shapes[i+1].Line.ForeColor.RGB = 0;
				slide.Shapes[i+1].Line.Weight = 5;
				slide.Shapes[i+1].Line.Visible = MsoTriState.msoCTrue;

			}
			float time = GetTimingsForImage(counter);
			totalTime += time;
			slide.SlideShowTransition.AdvanceTime = time;
			slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
			return totalTime;
		}

		private static float AddAnswerPage(Slides slides, int page, CustomLayout textLayout, Set set, int counter,
			float totalTime)
		{
			Slide slide;
			float time;
			slide = slides.AddSlide(page, textLayout);
			var title = slide.Shapes[1].TextFrame.TextRange;
			var text = set.IsValidSet() ? Tuple.Create("YES", 0x00CC00, "a Set") : Tuple.Create("NO", 0x0000FF, "not a Set");
			title.Text = text.Item1;
			title.Font.Name = "Arial Black";
			title.Font.Color.RGB = text.Item2;
			title.Font.Size = 80;
			var subtitle = slide.Shapes[2].TextFrame.TextRange;
			subtitle.Text = text.Item3;
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
			counter = counter%(count*2);
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
			counter = counter % (count * 2);
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