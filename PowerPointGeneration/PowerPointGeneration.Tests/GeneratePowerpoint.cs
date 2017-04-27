using System;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.SimpleLogger.Writers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PowerPointGeneration.Tests
{
    [TestClass]
    public class GeneratePowerpoint
    {
        public GeneratePowerpoint()
        {

            Logger.Writer = new ConsoleWriter();
           
        }
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
            CodeSmells.Create(new Details()
            {
                Name = "LongMethods",
                GoodName = "Short Enough",
                BadName = "Too Long",
                FontSize = 90
            });
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
        public void CreateSlidesForUserStories()
        {
            Logger.Writer = new ConsoleWriter();
            UserStoriesTraining.Create();
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
                BadName = "Nope",
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

           CreateSlidesForLongMethods();
            CodeSmells.Create(new Details()
            {
                Name = "LongLines",
                GoodName = "Short Enough",
                BadName = "Too Long",
                BackgroundColor = 0x272822,
                FontSize = 90
            });
            CodeSmells.Create(new Details()
            {
                Name = "BadNames",
                GoodName = "Good",
                BadName = "Bad",
            });
            CodeSmells.Create(new Details()
            {
                Name = "Clutter",
                GoodName = "Relevant",
                BadName = "Clutter",
            });
            CodeSmells.Create(new Details()
            {
                Name = "Duplication",
                GoodName = "Distinct",
                BadName = "Duplication",
                FontSize = 100
            });
            CodeSmells.Create(new Details()
            {
                Name = "Inconsistency",
                GoodName = "Duplication",
                BadName = "Inconsistency",
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
                BadName = "Too Long",
                BadNameText = "Pitkä",
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
                BadName = "Complex",
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
                BadName = "Go",
                Timings = slower
            });       
            CodeSmells.Create(new Details()
            {
                baseDirectory = @"C:\temp\languages\",
                FileNameFilter = "{1}.Assignment.{2:00}{3}",
                Name = "RustGo.Assignment",
                GoodName = "Rust",
                BadName = "Go",
                Timings = slower
            });
            CodeSmells.Create(new Details()
            {
                baseDirectory = @"C:\temp\languages\",
                FileNameFilter = "{1}.{2:00}{3}",
                Name = "RustHaskell",
                GoodName = "Rust",
                BadName = "Haskell",
                Timings = slower
            });
               CodeSmells.Create(new Details()
            {
                baseDirectory = @"C:\temp\languages\",
                FileNameFilter = "{1}.Assignment.{2:00}{3}",
                Name = "RustHaskell.Assignment",
                GoodName = "Rust",
                BadName = "Haskell",
                Timings = slower
            });
        }
    }
}