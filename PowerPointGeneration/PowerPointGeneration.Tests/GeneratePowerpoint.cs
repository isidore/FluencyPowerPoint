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

        [TestMethod]
        public void CreateSlidesForCodeSmells()
        {
            Logger.Writer = new ConsoleWriter();

            CodeSmells.Create(new Details()
            {
                Name = "LongMethods",
                GoodName = "Short Enough",
                GoodCount = 18,
                BadName = "Too Long",
                BadCount = 17,
                BackgroundColor = 0x272822,
                FontSize = 90
            });
            CodeSmells.Create(new Details()
            {
                Name = "BadNames",
                GoodName = "Good",
                GoodCount = 10,
                BadName = "Bad",
                BadCount = 29
            });
            CodeSmells.Create(new Details()
            {
                Name = "Clutter",
                GoodName = "Relevant",
                GoodCount = 28,
                BadName = "Clutter",
                BadCount = 28
            });
            CodeSmells.Create(new Details()
            {
                Name = "Duplication",
                GoodName = "Distinct",
                GoodCount = 4,
                BadName = "Duplication",
                BadCount = 7,
                FontSize = 100
            });
            CodeSmells.Create(new Details()
            {
                Name = "Inconsistency",
                GoodName = "Consistency",
                GoodCount = 1,
                BadName = "Inconsistency",
                BadCount = 4,
                FontSize = 100
            });
        }

        [TestMethod]
        public void CreateFinnishSmells()
        {
            CodeSmells.Create(new Details()
            {
                Name = "LongMethods",
                GoodName = "Short Enough",
                GoodNameText = "Lyhyt",
                GoodCount = 18,
                BadName = "Too Long",
                BadNameText = "Pitkä",
                BadCount = 17,
                BackgroundColor = 0x272822,
                FontSize = 90
            });
        }
    }
}