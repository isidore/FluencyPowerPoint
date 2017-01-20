using System;
using System.IO;
using System.Linq;
using ApprovalTests;
using ApprovalTests.Reporters;
using ApprovalUtilities.Utilities;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PowerPointGeneration.Tests
{
     [UseReporter(typeof(FileLauncherReporter))]
    [TestClass]
    public class FileRenamerTest
    {
        [TestMethod]
        public void Testname()
        {
            var name = @"C:\temp\birds\sparrow_song_";
            //TestRename(name);
            FileRenamer.Rename(name);
        }

         private static void TestRename(string name)
         {
             var reorder = FileRenamer.GetRenumbering(name);
             Approvals.VerifyAll(reorder, t => "{0} => {1}".FormatWith(t.Item1, t.Item2));
         }
    }

    public class FileRenamer
    {
        public static Tuple<string, string>[] GetRenumbering(string prefix)
        {
            var dir = Path.GetDirectoryName(prefix);
            var fileStart = Path.GetFileName(prefix);
            var files =  new DirectoryInfo(dir).EnumerateFiles(fileStart + "*");
            return files.Select((f,n) => Renumber(f,n,dir,fileStart)).ToArray();
        }

       

        private static Tuple<string, string> Renumber(FileInfo file, int number, string dir, string fileStart)
        {
            number ++;
            var extention = file.Extension;
            return Tuple.Create(file.FullName, @"{0}\{1}{2:00}{3}".FormatWith(dir, fileStart, number, extention));
        }

        public static void Rename(string name)
        {
            foreach (var r in GetRenumbering(name))
            {
                File.Move(r.Item1,r.Item2);
            }
        }
    }
}