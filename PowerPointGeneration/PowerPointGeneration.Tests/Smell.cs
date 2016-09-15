using ApprovalUtilities.Utilities;

namespace PowerPointGeneration.Tests
{
    public class Smell
    {
        public Details Details { get; set; }
        public bool Good { get; set; }

        public string fileName;

        public Smell(Details details, int number, bool good)
        {
            Details = details;
            Good = good;
            fileName = details.FileNameFilter.FormatWith(details.Name,
                good ? details.GoodName : details.BadName, number, details.FileEndingWithDot);
        }


        internal string GetImage()
        {
            return "{0}{1}".FormatWith(Details.baseDirectory, fileName);
        }
    }
}