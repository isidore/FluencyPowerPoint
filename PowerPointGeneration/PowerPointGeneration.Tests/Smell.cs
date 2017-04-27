using ApprovalUtilities.Utilities;

namespace PowerPointGeneration.Tests
{
    public class Smell
    {
        public Details Details { get; set; }
        public bool Good { get; set; }

        public string fileName;

        public Smell(Details details, int number, bool good):
            this(
            details, good, "{0}{1}".FormatWith(details.baseDirectory, details.FileNameFilter.FormatWith(details.Name,
                good ? details.GoodName : details.BadName, number, details.FileEndingWithDot) ))
        {
         
        }

        public Smell(Details details,  bool good, string fileName)
        {
            Details = details;
            Good = good;
            this.fileName = fileName;
        }


        internal string GetImage()
        {
            return fileName;
        }
    }
}