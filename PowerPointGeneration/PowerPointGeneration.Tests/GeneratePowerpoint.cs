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
                GoodCount = 15,
                BadName = "Too Long",
                BadCount = 26,
                BackgroundColor = 0x272822,
                FontSize = 90
            });
            CodeSmells.Create(new Details()
            {
                Name = "BadNames",
                GoodName = "Good",
                GoodCount = 27,
                BadName = "Bad",
                BadCount = 36
            });
            CodeSmells.Create(new Details()
            {
                Name = "Clutter",
                GoodName = "Relevant",
                GoodCount = 30,
                BadName = "Clutter",
                BadCount = 56
            });
            CodeSmells.Create(new Details()
            {
                Name = "Duplication",
                GoodName = "Distinct",
                GoodCount = 28,
                BadName = "Duplication",
                BadCount = 53,
                FontSize = 100
            });
            CodeSmells.Create(new Details()
            {
                Name = "Inconsistency",
                GoodName = "Duplication",
                GoodCount = 12,
                BadName = "Inconsistency",
                BadCount = 15,
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
        
        [TestMethod]
        public void CreateLanguageSlides()
        {
            Logger.Writer = new ConsoleWriter();
            var slower = new Timings {{2, 100}, {12, 4.0f}, {25, 3}, {Int32.MaxValue, 2}};
   
            CodeSmells.Create(new Details()
            {
                baseDirectory = @"C:\temp\languages\",
                FileNameFilter = "{1}.{2:00}{3}",
                Name = "RustGo",
                GoodName = "Rust",
                GoodCount = 23,
                BadName = "Go",
                BadCount = 23,
                Timings = slower
            });       
            CodeSmells.Create(new Details()
            {
                baseDirectory = @"C:\temp\languages\",
                FileNameFilter = "{1}.Assignment.{2:00}{3}",
                Name = "RustGo.Assignment",
                GoodName = "Rust",
                GoodCount = 18,
                BadName = "Go",
                BadCount = 19,
                Timings = slower
            });
            CodeSmells.Create(new Details()
            {
                baseDirectory = @"C:\temp\languages\",
                FileNameFilter = "{1}.{2:00}{3}",
                Name = "RustHaskell",
                GoodName = "Rust",
                GoodCount = 21,
                BadName = "Haskell",
                BadCount = 21,
                Timings = slower
            });
               CodeSmells.Create(new Details()
            {
                baseDirectory = @"C:\temp\languages\",
                FileNameFilter = "{1}.Assignment.{2:00}{3}",
                Name = "RustHaskell.Assignment",
                GoodName = "Rust",
                GoodCount = 18,
                BadName = "Haskell",
                BadCount = 17,
                Timings = slower
            });
        }
    }
}