using System.Linq;
using ApprovalUtilities.Utilities;

namespace PowerPointGeneration.Tests
{
	class Paragraph
	{
		public string fileName;
		private string number;

		const string BASE = @"C:\temp\Paragraphs\Paragraphs\";

		public Paragraph(string fileName)
		{
			// TODO: Complete member initialization
			this.fileName = fileName;
			this.number = fileName.Split('.').First();
		}

		internal string GetBasePicture()
		{
			return "{0}{1}.1.png".FormatWith(BASE, number);
		}

		internal string GetImage()
		{
			return "{0}{1}".FormatWith(BASE, fileName);
		}

		internal bool IsParagraph()
		{
			return fileName.Split('.')[2].Equals("yes");
		}
	}
}
