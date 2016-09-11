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

        [TestMethod]
        public void CreateSlidesForWellMaintained()
        {
            Logger.Writer = new ConsoleWriter();

            CodeSmells.Create(new Details()
            {
                Name = "WellMaintained",
                GoodName = "Yes",
                GoodCount = 4,
                BadName = "Nope",
                BadCount = 12,
                FileEndingWithDot = ".jpg",
                Timings = new Timings { { 2, 100 }, { 5, 2 }, { 20, 1.5F }, { Int32.MaxValue, 1 } }

            });
        } 
        
        [TestMethod]
        public void CreateSlidesForFunction()
        {
            Logger.Writer = new ConsoleWriter();
            FunctionalPrimer.Create();
        }

        [TestMethod]
        public void CreateSlidesForCodeSmells()
        {
            Logger.Writer = new ConsoleWriter();

            CodeSmells.Create(new Details()
            {
                Name = "LongMethods",
                GoodName = "Short Enough",
                GoodCount = 34,
                BadName = "Too Long",
                BadCount = 54,
                FontSize = 90
            });
            CodeSmells.Create(new Details()
            {
                Name = "LongLines",
                GoodName = "Short Enough",
                GoodCount = 8,
                BadName = "Too Long",
                BadCount = 14,
                BackgroundColor = 0x272822,
                FontSize = 90
            });
            CodeSmells.Create(new Details()
            {
                Name = "BadNames",
                GoodName = "Good",
                GoodCount = 12,
                BadName = "Bad",
                BadCount = 33
            });
            CodeSmells.Create(new Details()
            {
                Name = "Clutter",
                GoodName = "Relevant",
                GoodCount = 30,
                BadName = "Clutter",
                BadCount = 55
            });
            CodeSmells.Create(new Details()
            {
                Name = "Duplication",
                GoodName = "Distinct",
                GoodCount = 28,
                BadName = "Duplication",
                BadCount = 52,
                FontSize = 100
            });
            CodeSmells.Create(new Details()
            {
                Name = "Inconsistency",
                GoodName = "Duplication",
                GoodCount = 12,
                BadName = "Inconsistency",
                BadCount = 13,
                FontSize = 90,
                Timings = new Timings {{2, 100}, {5, 7}, {20, 5.5F}, {Int32.MaxValue, 4}}
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

        [TestMethod]
        public void CreateSlidesForCynefin()
        {
            Logger.Writer = new ConsoleWriter();

           
            CodeSmells.Create(new Details()
            {
                Name = "Complex",
                GoodName = "Complicated",
                GoodCount = 23,
                BadName = "Complex",
                BadCount = 21,
                FontSize = 90,
                Timings = new Timings { { 2, 100 }, { 12, 4.0f },  { 25, 3 }, { Int32.MaxValue, 2 } }
            });
        }
    }
}