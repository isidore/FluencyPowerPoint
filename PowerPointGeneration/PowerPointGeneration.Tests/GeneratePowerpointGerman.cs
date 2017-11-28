using System;
using ApprovalUtilities.SimpleLogger;
using ApprovalUtilities.SimpleLogger.Writers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PowerPointGeneration.Tests
{
    [TestClass]
    public class GeneratePowerpointGerman
    {
        public GeneratePowerpointGerman()
        {

            Logger.Writer = new ConsoleWriter();
           
        }
 
        [TestMethod]
        public void CreateSlidesForSparrows()
        {
            Logger.Writer = new ConsoleWriter();
            SparrowTraining.Create("Haus","Sing");
        }

        [TestMethod]
        public void CreateSlidesForLongMethods()
        {
            CodeSmells.Create(new Details()
            {
                Name = "LongMethods",
                GoodName = "Short Enough",
                GoodNameText = "Kurz",
                BadName = "Too Long",
                BadNameText = "Lang",
                FontSize = 90
            });
        }

       
        [TestMethod]
        public void CreateSlidesForWellMaintained()
        {
            Logger.Writer = new ConsoleWriter();

            CodeSmells.Create(new Details()
            {
                Name = "WellMaintained",
                GoodName = "Yes",
                GoodNameText = "Ja",
                BadName = "Nope",
                BadNameText = "Nein",
                FileEndingWithDot = ".jpg",
                Timings = new Timings { { 2, 100 }, { 5, 2 }, { 20, 1.5F }, { Int32.MaxValue, 1 } }

            });
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
                GoodNameText = "Kurz",
                BadName = "Too Long",
                BadNameText = "Lang",
                BackgroundColor = 0x272822,
                FontSize = 90
            });
            CodeSmells.Create(new Details()
            {
                Name = "BadNames",
                GoodName = "Good",
                GoodNameText = "Gut",
                BadName = "Bad",
                BadNameText = "Schlecht",
            });
            CodeSmells.Create(new Details()
            {
                Name = "Clutter",
                GoodName = "Relevant",
                GoodNameText = "Wertig",
                BadName = "Clutter",
                BadNameText = "Kram"
            });
            CodeSmells.Create(new Details()
            {
                Name = "Duplication",
                GoodName = "Distinct",
                GoodNameText = "Eindeutig",
                BadName = "Duplication",
                BadNameText = "Doppelung",
                FontSize = 100
            });

            CodeSmells.Create(new Details()
            {
                Name = "Inconsistency",
                GoodName = "Duplication",
                GoodNameText = "Doppelung",
                BadName = "Inconsistency",
                BadNameText = "Inkonsistenz",
                FontSize = 90,
                Timings = new Timings {{2, 100}, {5, 7}, {20, 5.5F}, {Int32.MaxValue, 4}}
            });
        }

       
    }
}