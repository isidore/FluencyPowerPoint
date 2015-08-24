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
	public class CodeParagraphsTrainings
	{
		public static void Create()
		{
			Application pptApplication = new Application();
			// Create the Presentation File
			Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
			AddCode(pptPresentation);
			pptPresentation.SaveAs(@"c:\temp\CodeParagraphs.pptx", PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
			pptPresentation.Close();
		}

		private static Paragraph[] GetTrainingSet()
		{
			var files = GetFiles();
			return files.Select(f => new Paragraph(f)).Shuffle().ToArray().Log("# of Examples", v => "" + v.Count());
		}

		private static string[] GetFiles()
		{
			var hard = new[] {
"01.2.yes.png",
"01.3.no.png",
"02.2.yes.png"
,"02.3.no.png"
,"03.2.no.png"
,"03.2.yes.png"
,"04.2.yes.png"
,"04.3.yes.png"
,"04.4.no.png"
,"05.2.yes.png"
,"05.3.no.png"
,"06.2.yes.png"
,"06.3.yes.png"
,"06.4.no.png"
,"06.5.no.png"
,"07.2.yes.png"
,"07.3.yes.png"
,"07.4.no.png"
,"07.5.no.png"
,"08.2.no.png"
,"08.2.yes.png"
,"08.3.no.png"
,"09.1.yes.png"
,"09.2.no.png"
,"10.2.yes.png"
,"10.3.no.png"
,"10.4.no.png" };
			var system = Enumerable.Range(11, 20).SelectMany(n => new[] { n + ".2.yes.png", n + ".3.no.png" });
			return hard.Concat(system).ToArray();
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
					page+= 2;

			
					
					// Answer
					totalTime = AddAnswerPage(slides, page, textLayout, code, counter, totalTime);
					page+=1;
				}
				Logger.Variable("Total Time", "{0:00}:{0:00}".FormatWith(totalTime/60, totalTime%60));
			}
		}

		private static float AddAnswerPage(Slides slides, int page, CustomLayout customLayout,
			Paragraph code, int counter, float totalTime)
		{
			return AddAnswerImage(slides, page, customLayout, totalTime, 0.5f, code);
		}

		private static float AddPicturePage(Slides slides, int page, CustomLayout customLayout,
			Paragraph code, int counter, float totalTime)
		{
			float time = GetTimingsForImage(counter);
			var image = code.GetBasePicture();
			totalTime = AddImage(slides, page, customLayout, totalTime, 0.5f, image);
			return AddImage(slides, page + 1, customLayout, totalTime, time, code.GetImage());
		}

		private static float AddAnswerImage(Slides slides, int page, CustomLayout customLayout,  float totalTime, float time, Paragraph paragraph)
		{
			var slide = slides.AddSlide(page, customLayout);
			Shape shape = slide.Shapes[2];
			shape.Top = 0;
			shape.Left = 0;
			shape.Height = slide.Design.SlideMaster.Height;
			shape.Width = slide.Design.SlideMaster.Width;
			slide.Shapes.AddPicture(paragraph.GetImage(), MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top,
				shape.Width, shape.Height);
			slide.Background.Fill.ForeColor.RGB = 0xFFFFFF;
			slide.FollowMasterBackground = MsoTriState.msoFalse;
			var title = slide.Shapes[1].TextFrame.TextRange;
			title.Text = paragraph.IsParagraph() ? "Paragraph" : "No";
			title.Font.Name = "Arial Black";
			title.Font.Size = 120;
			slide.Shapes[1].Top = 0;
			slide.Shapes[1].Left = 0;
			slide.Shapes[1].Width = slide.Design.SlideMaster.Width;
			var color = paragraph.IsParagraph() ? 0x347400 : 0x3B3BFF;
			title.Font.Color.RGB = color;
			slide.Shapes[1].ZOrder(MsoZOrderCmd.msoBringToFront);
			totalTime += time;
			slide.SlideShowTransition.AdvanceTime = time;
			slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
			slide.NotesPage.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 0, 0);
			var t = slide.NotesPage.Shapes[2];
      t.TextFrame.TextRange.Text =  paragraph.fileName;
			
			return totalTime;
		}

		public static float AddImage(Slides slides, int page, CustomLayout customLayout, float totalTime, float time, string image)
		{
			var slide = slides.AddSlide(page, customLayout);
			Shape shape = slide.Shapes[2];
			shape.Top = 0;
			shape.Left = 0;
			shape.Height = slide.Design.SlideMaster.Height;
			shape.Width = slide.Design.SlideMaster.Width;
			slide.Shapes.AddPicture(image, MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top,
				shape.Width, shape.Height);
			slide.Background.Fill.ForeColor.RGB = 0xFFFFFF;
			slide.FollowMasterBackground = MsoTriState.msoFalse;
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
				return 3f;
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

	public static class LogUtils
	{
		public static T Log<T>(this T t, string label, Func<T, string> log)
		{
			Logger.Variable(label, log(t));
			return t;
		}
	}
}