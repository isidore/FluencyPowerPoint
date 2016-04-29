using ApprovalUtilities.Utilities;

namespace PowerPointGeneration.Tests
{
    public class Smell
    {
        public Details Details { get; set; }
        public bool Good { get; set; }

        public string fileName;
        private const string BASE = @"C:\code\FluencyPowerPoint\PowerPointGeneration\PowerPointGeneration.Tests\";

        public Smell(Details details, int number, bool good)
        {
            Details = details;
            Good = good;
            fileName = "CodeSmells-{0}\\{1} {2:00}{3}".FormatWith(details.Name,
                good ? details.GoodName : details.BadName, number, details.FileEndingWithDot);
        }


        internal string GetImage()
        {
            return "{0}{1}".FormatWith(BASE, fileName);
        }
    }
}