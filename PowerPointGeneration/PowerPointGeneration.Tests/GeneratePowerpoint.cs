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
	}
}
