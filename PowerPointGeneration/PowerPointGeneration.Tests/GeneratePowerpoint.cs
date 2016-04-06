﻿using System;
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
                GoodCount = 7,
                BadName = "Nope",
                BadCount = 12,
                FileEndingWithDot = ".jpg"
            });
        }

        [TestMethod]
        public void CreateSlidesForCodeSmells()
        {
            Logger.Writer = new ConsoleWriter();

            CodeSmells.Create(new Details()
            {
                Name = "LongMethods",
                GoodName = "Short Enough",
                GoodCount = 35,
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
                GoodCount = 4,
                BadName = "Duplication",
                BadCount = 52,
                FontSize = 100
            });
            CodeSmells.Create(new Details()
            {
                Name = "Inconsistency",
                GoodName = "Consistency",
                GoodCount = 1,
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
    }
}