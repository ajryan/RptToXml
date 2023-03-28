using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace RptToXml
{
	class Program
	{
		static void Main(string[] args)
		{
			if (args.Length < 1)
			{
				Console.WriteLine("Usage: RptToXml.exe < -r | RPT filename | wildcard> [outputfilename] [--ignore-errors]");
				Console.WriteLine("       -r : recursively convert all rpt files in current directory and sub directories");
				Console.WriteLine("       outputfilename argument is valid only with single filename in first argument");
				return;
			}
			Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
			string rptPathArg = args[0];
			var rptPaths = new List<string>();
			int ignoreErrFlag = 0;

			if (args.Contains("--ignore-errors"))
				ignoreErrFlag = 1;

			if ("-r".Equals(rptPathArg, StringComparison.InvariantCultureIgnoreCase))
			{
				if (args.Length > 1 + ignoreErrFlag)
				{
					Console.WriteLine("Output filename may not be specified with -r .");
					return;
				}
				recursiveFileList(rptPaths, ".");
			}
			else if (rptPathArg.Contains("*"))
			{
				if (args.Length > 1 + ignoreErrFlag)
				{
					Console.WriteLine("Output filename may not be specified with wildcard.");
					return;
				}
				var directory = Path.GetDirectoryName(rptPathArg);
				if (String.IsNullOrEmpty(directory))
				{
					directory = ".";
				}
				var matchingFiles = Directory.GetFiles(directory, searchPattern: Path.GetFileName(rptPathArg));
				rptPaths.AddRange(matchingFiles.Where(ReportFilenameValid));
				if (rptPaths.Count == 0)
				{
					Trace.WriteLine("No reports matched the wildcard.");
				}
			}
			else
			{
				rptPaths.Add(rptPathArg);
			}

			foreach (string rptPath in rptPaths)
			{
				string xmlPath = args.Length > 1 + ignoreErrFlag ?
					args[1] : Path.ChangeExtension(rptPath, "xml");
				DumpFile(rptPath, xmlPath, ignoreErrFlag == 1);
			}
		}
		static void DumpFile(string rptPath, string xmlPath, bool ignoreErr)
		{
			try
			{
				Trace.WriteLine("Dumping " + rptPath);

				using (var writer = new RptDefinitionWriter(rptPath))
				{
					writer.WriteToXml(xmlPath);
				}
			}
			catch (Exception ex)
			{
				if (ignoreErr)
					Trace.WriteLine(ex.Message);
				else
					throw ex;
			}
		}
		static void recursiveFileList(List<string> list, string directory)
		{
			foreach (string f in Directory.GetFiles(directory, "*.rpt"))
			{
				list.Add(f);
			}
			foreach (string d in Directory.GetDirectories(directory))
			{
				recursiveFileList(list, d);
			}
		}
		static bool ReportFilenameValid(string rptPath)
		{
			string extension = Path.GetExtension(rptPath);
			if (String.IsNullOrEmpty(extension) || !extension.Equals(".rpt", StringComparison.OrdinalIgnoreCase))
			{
				Console.WriteLine("Input filename [" + rptPath + "] does not end in .RPT");
				return false;
			}

			if (!File.Exists(rptPath))
			{
				Console.WriteLine("Report file [" + rptPath + "] does not exist.");
				return false;
			}

			return true;
		}
	}
}
