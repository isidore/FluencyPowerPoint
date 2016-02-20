using System;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.SimpleLogger.Writers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PowerPointGeneration.Tests
{
	[TestClass]
	public class GeneratePowerpoint
	{
		[TestMethod]
		public void CreateSlidesForSparrows()
		{
			Logger.Writer = new ConsoleWriter();
			SparrowTraining.Create();
		}
		[TestMethod]
		public void CreateSlidesForSet()
		{
			Logger.Writer = new ConsoleWriter();
			SetTraining.Create();
		}
		[TestMethod]
		public void CreateSlidesForLongMethods()
		{
			Logger.Writer = new ConsoleWriter();
			LongMethodsTraining.Create();
		}
		[TestMethod]
		public void CreateSlidesForUnitTestStories()
		{
			Logger.Writer = new ConsoleWriter();
			UnitTestStoryTraining.Create();
		}

		[TestMethod]
		public void CreateSlidesForCodeParagraphs()
		{
			Logger.Writer = new ConsoleWriter();
			CodeParagraphsTrainings.Create();
		}
        
        [TestMethod]
		public void CreateSlidesForPropertyBasedTesting()
		{
			Logger.Writer = new ConsoleWriter();
            PropertyBasedTestingTraining.Create();
		}

[TestMethod]
        public void CreateSlidesForCodeSmellsBadNames()
		{
			Logger.Writer = new ConsoleWriter();
            CodeSmellsBadNames.Create();
		}


	}
}
