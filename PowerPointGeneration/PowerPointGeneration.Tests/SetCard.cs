using System;
using ApprovalUtilities.Utilities;

namespace PowerPointGeneration.Tests
{
	public class SetCard
	{
		public const int SWIGGLE = 1;
		public const int DIAMOND = 2;
		public const int OVAL = 3;
		public const int RED = 1;
		public const int PURPLE = 2;
		public const int GREEN = 3;
		public const int SOLID = 1;
		public const int LINED = 2;
		public const int EMPTY = 3;
		public int symbol = 0;
		public int number = 0;
		public int shading = 0;
		public int color = 0;

		public SetCard(int graphicNumber)
		{
			if ((graphicNumber < 1) || (graphicNumber > 81))
			{
				throw new Exception("SetCard:Tried to create card with graphicNumber(" + graphicNumber + ")");
			}

			graphicNumber--;
			shading = ((graphicNumber%81)/27) + 1;
			symbol = ((graphicNumber%27)/9) + 1;
			color = ((graphicNumber%9)/3) + 1;
			number = ((graphicNumber%3)/1) + 1;
		}

		
		public static string GetImageFileName(int cardNumber)
		{
			return @"c:\temp\setcards\card{0:00}.gif".FormatWith(cardNumber);
		}

		public int GetGraphicsNumber()
		{
			var num = (shading - 1)*27;
			num += (symbol - 1)*9;
			num += (color - 1)*3;
			num += number; //should be (Number - 1) * 1 but over all sum needs to be inc by one anyhow...
			return num;
		}


		public override string ToString()
		{
			return "[Shading = " + shading + ",Symbol = " + symbol +
			       ",Color = " + color + ",Number = " + number + "]";
		}

		public string GetImageFileName()
		{
			return GetImageFileName(GetGraphicsNumber());
		}
	}
}