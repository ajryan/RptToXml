using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace RptToXml
{
	class Program
	{
		static void Main(string[] args)
		{
			if (args.Length != 1)
			{
				Console.WriteLine("Usage: RptToXml.exe <RPT filename | wildcard>");
				return;
			}

			string rptPathArg = args[0];
			bool wildCard = rptPathArg.Contains("*");
			if (!wildCard && !ValidateReportFilename(rptPathArg))
				return;

			Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));

			var rptPaths = new List<string>();
			if (!wildCard)
			{
				rptPaths.Add(rptPathArg);
			}
			else
			{
				var matchingFiles = Directory.GetFiles(Path.GetDirectoryName(rptPathArg), Path.GetFileName(rptPathArg));
				rptPaths.AddRange(matchingFiles);
			}

			foreach (string rptPath in rptPaths)
			{
				Trace.WriteLine("Dumping " + rptPath);

				RptDefinitionWriter writer = new RptDefinitionWriter(rptPath);

				string xmlPath = Path.ChangeExtension(rptPath, "xml");
				writer.WriteToXml(xmlPath);
			}
		}

		static bool ValidateReportFilename(string rptPath)
		{
			if (!Path.GetExtension(rptPath).Equals(".rpt", StringComparison.OrdinalIgnoreCase))
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
