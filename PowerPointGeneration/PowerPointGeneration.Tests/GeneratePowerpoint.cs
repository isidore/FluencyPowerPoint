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
		public void CreateSlides()
		{
			Logger.Writer = new ConsoleWriter();
			Sparrows.Create();
		}
	}
}
