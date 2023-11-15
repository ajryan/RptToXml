using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Runtime.ExceptionServices;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportAppServer.ClientDoc;
using CrystalDecisions.ReportAppServer.Controllers;
using CrystalDecisions.Shared;

using CRDataDefModel = CrystalDecisions.ReportAppServer.DataDefModel;
using CRReportDefModel = CrystalDecisions.ReportAppServer.ReportDefModel;

using OpenMcdf;

namespace RptToXml
{
	public partial class RptDefinitionWriter : IDisposable
	{
		private const FormatTypes ShowFormatTypes = FormatTypes.AreaFormat | FormatTypes.SectionFormat | FormatTypes.Color;

		private ReportDocument _report;
		private ISCDReportClientDocument _rcd;
		private CompoundFile _oleCompoundFile;

		private readonly bool _createdReport;
		private readonly bool _stdOut;

		public RptDefinitionWriter(string filename, bool stdOut)
		{
			_stdOut = stdOut;
			_createdReport = true;
			_report = new ReportDocument();
			_report.Load(filename, OpenReportMethod.OpenReportByTempCopy);
			_rcd = _report.ReportClientDocument;

			_oleCompoundFile = new CompoundFile(filename);

			if (!stdOut)
            {
                Trace.WriteLine("Loaded report");
            }
        }

		public RptDefinitionWriter(ReportDocument value)
		{
			_report = value;
			_rcd = _report.ReportClientDocument;
		}

		public void WriteToXml()
		{
			XmlWriterSettings settings = new XmlWriterSettings
            {
                CheckCharacters = true,
                Encoding = Encoding.UTF8,
                Indent = true
            };

            StringBuilder stringOutput = new StringBuilder();
			using (XmlWriter writer = XmlWriter.Create(stringOutput, settings))
			{
				WriteToXml(writer);
			}
			Trace.Write(stringOutput.ToString());
		}

        public void WriteToXml(string targetXmlPath)
        {
            WriteToXml(System.IO.File.Create(targetXmlPath));
        }

		public void WriteToXml(System.IO.Stream output)
		{

			XmlWriterSettings settings = new XmlWriterSettings
            {
                CheckCharacters = true,
                Encoding = Encoding.UTF8,
                Indent = true
            };
            using (XmlWriter writer = XmlWriter.Create(output, settings))
			{
				WriteToXml(writer);
			}
		}

		public void WriteToXml(XmlWriter writer)
		{
			if (!_stdOut)
            {
                Trace.WriteLine("Writing to XML");
            }

            writer.WriteStartDocument();
			ProcessReport(_report, writer);
			writer.WriteEndDocument();
			writer.Flush();
		}

		private static readonly System.Text.RegularExpressions.Regex CompiledRegexp =
			new System.Text.RegularExpressions.Regex("[^\\u0009\\u000a\\u000d\\u0020-\\uD7FF\\uE000-\\uFFFD]",
				System.Text.RegularExpressions.RegexOptions.Compiled);

		private static void WriteAttributeString(XmlWriter writer, string name, string value)
		{
			string myValue = value == null ? null : CompiledRegexp.Replace(value, "");
			writer.WriteAttributeString(name, myValue ?? "");
		}
		private static void WriteString(XmlWriter writer, string value)
		{
			string myValue = value == null ? null : CompiledRegexp.Replace(value, "");
			writer.WriteString(myValue ?? "");
		}

        //This is a recursive method.  GetSubreports() calls it.
		private void ProcessReport(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("Report");

			WriteAttributeString(writer, "Name", report.Name);
			if (!_stdOut)
            {
                Trace.WriteLine("Writing report " + report.Name);
            }

            if (!report.IsSubreport)
			{
				if (!_stdOut)
                {
                    Trace.WriteLine("Writing header info");
                }

                WriteAttributeString(writer, "FileName", report.FileName.Replace("rassdk://", ""));
				WriteAttributeString(writer, "HasSavedData", report.HasSavedData.ToString());

				if (_oleCompoundFile != null)
				{
					writer.WriteStartElement("Embedinfo");
					_oleCompoundFile.RootStorage.VisitEntries(fileItem =>
					{
						if (fileItem.Name.Contains("Ole"))
						{
							writer.WriteStartElement("Embed");
							WriteAttributeString(writer, "Name", fileItem.Name);

							var cfStream = fileItem as CFStream;
							if (cfStream != null)
							{
								var streamBytes = cfStream.GetData();

								WriteAttributeString(writer, "Size", cfStream.Size.ToString("0"));

								using (var md5Provider = new MD5CryptoServiceProvider())
								{
									byte[] md5Hash = md5Provider.ComputeHash(streamBytes);
									WriteAttributeString(writer, "MD5Hash", Convert.ToBase64String(md5Hash));
								}
							}
							writer.WriteEndElement();
						}
					}, true);
					writer.WriteEndElement();
				}

				GetSummaryinfo(report, writer);
				GetReportOptions(report, writer);
				GetPrintOptions(report, writer);
				GetSubreports(report, writer);  //recursion happens here.
			}

			GetDatabase(report, writer);
			GetDataDefinition(report, writer);
			GetCustomFunctions(report, writer);
			GetSubReportsLinks(report, writer);
			GetReportDefinition(report, writer);

			writer.WriteEndElement();
		}

		private static void GetSummaryinfo(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("Summaryinfo");

			WriteAttributeString(writer, "KeywordsinReport", report.SummaryInfo.KeywordsInReport);
			WriteAttributeString(writer, "ReportAuthor", report.SummaryInfo.ReportAuthor);
			WriteAttributeString(writer, "ReportComments", report.SummaryInfo.ReportComments);
			WriteAttributeString(writer, "ReportSubject", report.SummaryInfo.ReportSubject);
			WriteAttributeString(writer, "ReportTitle", report.SummaryInfo.ReportTitle);

			writer.WriteEndElement();
		}

		private static void GetReportOptions(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("ReportOptions");

			WriteAttributeString(writer, "EnableSaveDataWithReport", report.ReportOptions.EnableSaveDataWithReport.ToString());
			WriteAttributeString(writer, "EnableSavePreviewPicture", report.ReportOptions.EnableSavePreviewPicture.ToString());
			WriteAttributeString(writer, "EnableSaveSummariesWithReport", report.ReportOptions.EnableSaveSummariesWithReport.ToString());
			WriteAttributeString(writer, "EnableUseDummyData", report.ReportOptions.EnableUseDummyData.ToString());
			WriteAttributeString(writer, "initialDataContext", report.ReportOptions.InitialDataContext);
			WriteAttributeString(writer, "initialReportPartName", report.ReportOptions.InitialReportPartName);

			writer.WriteEndElement();
		}

		private void GetPrintOptions(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("PrintOptions");

			WriteAttributeString(writer, "PageContentHeight", report.PrintOptions.PageContentHeight.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "PageContentWidth", report.PrintOptions.PageContentWidth.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "PaperOrientation", report.PrintOptions.PaperOrientation.ToString());
			WriteAttributeString(writer, "PaperSize", report.PrintOptions.PaperSize.ToString());
			WriteAttributeString(writer, "PaperSource", report.PrintOptions.PaperSource.ToString());
			WriteAttributeString(writer, "PrinterDuplex", report.PrintOptions.PrinterDuplex.ToString());
			WriteAttributeString(writer, "PrinterName", report.PrintOptions.PrinterName);

			writer.WriteStartElement("PageMargins");

			WriteAttributeString(writer, "bottomMargin", report.PrintOptions.PageMargins.bottomMargin.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "leftMargin", report.PrintOptions.PageMargins.leftMargin.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "rightMargin", report.PrintOptions.PageMargins.rightMargin.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "topMargin", report.PrintOptions.PageMargins.topMargin.ToString(CultureInfo.InvariantCulture));

			writer.WriteEndElement();

			CRReportDefModel.PrintOptions rdmPrintOptions = GetRASRDMPrintOptionsObject(report.Name, report);
			if (rdmPrintOptions != null)
            {
                GetPageMarginConditionFormulas(rdmPrintOptions, writer);
            }

            writer.WriteEndElement();
		}

		private void GetSubReportsLinks(ReportDocument report, XmlWriter writer)
		{
			if (report.IsSubreport)
			{
				writer.WriteStartElement("SubReportLinks");
				CRReportDefModel.SubreportLinks subReportLinks = _report.ReportClientDocument.SubreportController.GetSubreportLinks(report.Name);

				if (subReportLinks != null)
                {
                    foreach (CRReportDefModel.SubreportLink link in subReportLinks)
                    {
                        writer.WriteStartElement("SubReportLink");
                        WriteAttributeString(writer, "LinkedParameterName", link.LinkedParameterName);
                        WriteAttributeString(writer, "MainReportFieldName", link.MainReportFieldName);
                        WriteAttributeString(writer, "SubreportFieldName", link.SubreportFieldName);
                        writer.WriteEndElement();
                    }
                }

                writer.WriteEndElement();
			}
		}

		[HandleProcessCorruptedStateExceptions]
		private void GetSubreports(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("SubReports");

			try
			{
				foreach (ReportDocument subreport in report.Subreports)
                {
                    ProcessReport(subreport, writer);
                }
            }
			catch (Exception e)
			{
				Console.WriteLine($"Error loading subpreport, {e}");
			}
			writer.WriteEndElement();
		}

		private void GetDatabase(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("Database");

			GetTableLinks(report, writer);
			if (!report.IsSubreport)
			{
				var reportClientDocument = report.ReportClientDocument;
				GetReportClientTables(reportClientDocument, writer);
			}
			else
			{
				var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
				GetSubreportClientTables(subrptClientDoc, writer);
			}

			writer.WriteEndElement();
		}

		private static void GetTableLinks(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("TableLinks");

			foreach (TableLink tl in report.Database.Links)
			{
				writer.WriteStartElement("TableLink");
				WriteAttributeString(writer, "JoinType", tl.JoinType.ToString());

				writer.WriteStartElement("SourceFields");
				foreach (FieldDefinition fd in tl.SourceFields)
                {
                    GetFieldDefinition(fd, writer);
                }

                writer.WriteEndElement();

				writer.WriteStartElement("DestinationFields");
				foreach (FieldDefinition fd in tl.DestinationFields)
                {
                    GetFieldDefinition(fd, writer);
                }

                writer.WriteEndElement();

				writer.WriteEndElement();
			}

			writer.WriteEndElement();
		}

		private void GetCustomFunctions(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("CustomFunctions");

            // TODO:
            // var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
            // funcs = subrptClientDoc.CustomFunctionController.GetCustomFunctions();
            CRDataDefModel.CustomFunctions funcs = !report.IsSubreport
                ? report.ReportClientDocument.CustomFunctionController.GetCustomFunctions()
                : null;

			if (funcs != null)
			{
				foreach (CRDataDefModel.CustomFunction func in funcs)
				{
					writer.WriteStartElement("CustomFunction");
					WriteAttributeString(writer, "Name", func.Name);
					WriteAttributeString(writer, "Syntax", func.Syntax.ToString());
					writer.WriteElementString("Text", func.Text); // an element so line breaks are literal

					writer.WriteEndElement();
				}
			}

			writer.WriteEndElement();
		}

		private void GetReportClientTables(ISCDReportClientDocument reportClientDocument, XmlWriter writer)
		{
			writer.WriteStartElement("Tables");

			foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table table in reportClientDocument.DatabaseController.Database.Tables)
			{
				GetTable(table, writer);
			}

			writer.WriteEndElement();
		}
		private void GetSubreportClientTables(SubreportClientDocument subrptClientDocument, XmlWriter writer)
		{
			writer.WriteStartElement("Tables");

			foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table table in subrptClientDocument.DatabaseController.Database.Tables)
			{
				GetTable(table, writer);
			}

			writer.WriteEndElement();
		}

		private void GetTable(CrystalDecisions.ReportAppServer.DataDefModel.Table table, XmlWriter writer)
		{
			writer.WriteStartElement("Table");

			WriteAttributeString(writer, "Alias", table.Alias);
			WriteAttributeString(writer, "ClassName", table.ClassName);
			WriteAttributeString(writer, "Name", table.Name);

			writer.WriteStartElement("ConnectionInfo");
			foreach (string propertyId in table.ConnectionInfo.Attributes.PropertyIDs)
			{
				// make attribute name safe for XML
				string attributeName = propertyId.Replace(" ", "_");

				WriteAttributeString(writer, attributeName, table.ConnectionInfo.Attributes[propertyId].ToString());
			}

			WriteAttributeString(writer, "UserName", table.ConnectionInfo.UserName);
			WriteAttributeString(writer, "Password", table.ConnectionInfo.Password);
			writer.WriteEndElement();

            if (table is CRDataDefModel.CommandTable commandTable)
			{
				var cmdTable = commandTable;
				writer.WriteStartElement("Command");
				WriteString(writer, cmdTable.CommandText);
				writer.WriteEndElement();
			}

			writer.WriteStartElement("Fields");

			foreach (CrystalDecisions.ReportAppServer.DataDefModel.Field fd in table.DataFields)
			{
				GetFieldDefinition(fd, writer);
			}

			writer.WriteEndElement();

			writer.WriteEndElement();
		}

		private static void GetFieldDefinition(FieldDefinition fd, XmlWriter writer)
		{
			writer.WriteStartElement("Field");

			WriteAttributeString(writer, "FormulaName", fd.FormulaName);
			WriteAttributeString(writer, "Kind", fd.Kind.ToString());
			WriteAttributeString(writer, "Name", fd.Name);
			WriteAttributeString(writer, "NumberOfBytes", fd.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "ValueType", fd.ValueType.ToString());

			writer.WriteEndElement();
		}

		private static void GetFieldDefinition(CrystalDecisions.ReportAppServer.DataDefModel.Field fd, XmlWriter writer)
		{
			writer.WriteStartElement("Field");

			WriteAttributeString(writer, "Description", fd.Description);
			WriteAttributeString(writer, "FormulaForm", fd.FormulaForm);
			WriteAttributeString(writer, "HeadingText", fd.HeadingText);
			WriteAttributeString(writer, "IsRecurring", fd.IsRecurring.ToString());
			WriteAttributeString(writer, "Kind", fd.Kind.ToString());
			WriteAttributeString(writer, "Length", fd.Length.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "LongName", fd.LongName);
			WriteAttributeString(writer, "Name", fd.Name);
			WriteAttributeString(writer, "ShortName", fd.ShortName);
			WriteAttributeString(writer, "Type", fd.Type.ToString());
			WriteAttributeString(writer, "UseCount", fd.UseCount.ToString(CultureInfo.InvariantCulture));

			writer.WriteEndElement();
		}

		private void GetDataDefinition(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("DataDefinition");

			writer.WriteElementString("GroupSelectionFormula", report.DataDefinition.GroupSelectionFormula);
			writer.WriteElementString("RecordSelectionFormula", report.DataDefinition.RecordSelectionFormula);

			writer.WriteStartElement("Groups");
			foreach (Group group in report.DataDefinition.Groups)
			{
				writer.WriteStartElement("Group");
				WriteAttributeString(writer, "ConditionField", group.ConditionField.FormulaName);

				writer.WriteEndElement();

			}
			writer.WriteEndElement();

			writer.WriteStartElement("SortFields");
			foreach (SortField sortField in report.DataDefinition.SortFields)
			{
				writer.WriteStartElement("SortField");

				WriteAttributeString(writer, "Field", sortField.Field.FormulaName);
				try
				{
					string sortDirection = sortField.SortDirection.ToString();
					WriteAttributeString(writer, "SortDirection", sortDirection);
				}
				catch (NotSupportedException)
				{ }
				WriteAttributeString(writer, "SortType", sortField.SortType.ToString());

				writer.WriteEndElement();
			}
			writer.WriteEndElement();

			writer.WriteStartElement("FormulaFieldDefinitions");
			foreach (var field in report.DataDefinition.FormulaFields.OfType<FieldDefinition>().OrderBy(field => field.FormulaName))
            {
                GetFieldObject(field, report, writer);
            }

            writer.WriteEndElement();

			writer.WriteStartElement("GroupNameFieldDefinitions");
			foreach (var field in report.DataDefinition.GroupNameFields)
            {
                GetFieldObject(field, report, writer);
            }

            writer.WriteEndElement();

			writer.WriteStartElement("ParameterFieldDefinitions");
			try
			{
				foreach (var field in report.DataDefinition.ParameterFields)
                {
                    GetFieldObject(field, report, writer);
                }
            }
			catch (Exception e)
			{
				Console.WriteLine($"Error processing ParameterFieldDefinitions, {e}");
			}
			writer.WriteEndElement();

			writer.WriteStartElement("RunningTotalFieldDefinitions");
			foreach (var field in report.DataDefinition.RunningTotalFields)
            {
                GetFieldObject(field, report, writer);
            }

            writer.WriteEndElement();

			writer.WriteStartElement("SQLExpressionFields");
			foreach (var field in report.DataDefinition.SQLExpressionFields)
            {
                GetFieldObject(field, report, writer);
            }

            writer.WriteEndElement();

			writer.WriteStartElement("SummaryFields");
			foreach (var field in report.DataDefinition.SummaryFields)
            {
                GetFieldObject(field, report, writer);
            }

            writer.WriteEndElement();

			writer.WriteEndElement();
		}

		[HandleProcessCorruptedStateExceptions]
		private void GetFieldObject(Object fo, ReportDocument report, XmlWriter writer)
		{
			if (fo is DatabaseFieldDefinition df)
			{
                writer.WriteStartElement("DatabaseFieldDefinition");

				WriteAttributeString(writer, "FormulaName", df.FormulaName);
				WriteAttributeString(writer, "Kind", df.Kind.ToString());
				WriteAttributeString(writer, "Name", df.Name);
				WriteAttributeString(writer, "NumberOfBytes", df.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "TableName", df.TableName);
				WriteAttributeString(writer, "ValueType", df.ValueType.ToString());

			}
			else if (fo is FormulaFieldDefinition ff)
			{
				var ddm_ff = GetRASDDMFormulaFieldObject(ff.Name, report);
			

				writer.WriteStartElement("FormulaFieldDefinition");

				WriteAttributeString(writer, "FormulaName", ff.FormulaName);
				WriteAttributeString(writer, "Kind", ff.Kind.ToString());
				WriteAttributeString(writer, "Name", ff.Name);
				WriteAttributeString(writer, "NumberOfBytes", ff.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "ValueType", ff.ValueType.ToString());
				if (ddm_ff != null)
				{
					WriteAttributeString(writer, "Syntax", ddm_ff.Syntax.ToString());
				}
				WriteString(writer, ff.Text);

			}
			else if (fo is GroupNameFieldDefinition gnf)
			{
                writer.WriteStartElement("GroupNameFieldDefinition");
				try
				{
					WriteAttributeString(writer, "FormulaName", gnf.FormulaName);
					WriteAttributeString(writer, "Group", gnf.Group.ToString());
					WriteAttributeString(writer, "GroupNameFieldName", gnf.GroupNameFieldName);
					WriteAttributeString(writer, "Kind", gnf.Kind.ToString());
					WriteAttributeString(writer, "Name", gnf.Name);
					WriteAttributeString(writer, "NumberOfBytes", gnf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
					WriteAttributeString(writer, "ValueType", gnf.ValueType.ToString());
				}
				catch (Exception e)
				{
					Console.WriteLine($"Error loading formula for group '{gnf.GroupNameFieldName}', {e}");
				}
			}
			else if (fo is ParameterFieldDefinition pf)
			{
                // if it is a linked parameter, it is passed into a subreport. Just record the actual linkage in the main report.
				// The parameter will be reported in full when the subreport is exported.  
				var parameterIsLinked = (!report.IsSubreport && pf.IsLinked());

				writer.WriteStartElement("ParameterFieldDefinition");

				if (parameterIsLinked)
				{
					WriteAttributeString(writer, "Name", pf.Name);
					WriteAttributeString(writer, "IsLinkedToSubreport", pf.IsLinked().ToString());
					WriteAttributeString(writer, "ReportName", pf.ReportName);
				}
				else
				{
					var ddm_pf = GetRASDDMParameterFieldObject(pf.Name, report);

					WriteAttributeString(writer, "AllowCustomCurrentValues", (ddm_pf != null && ddm_pf.AllowCustomCurrentValues).ToString());
					WriteAttributeString(writer, "EditMask", pf.EditMask);
					WriteAttributeString(writer, "EnableAllowEditingDefaultValue", pf.EnableAllowEditingDefaultValue.ToString());
					WriteAttributeString(writer, "EnableAllowMultipleValue", pf.EnableAllowMultipleValue.ToString());
					WriteAttributeString(writer, "EnableNullValue", pf.EnableNullValue.ToString());
					WriteAttributeString(writer, "FormulaName", pf.FormulaName);
					WriteAttributeString(writer, "HasCurrentValue", pf.HasCurrentValue.ToString());
					WriteAttributeString(writer, "IsOptionalPrompt", pf.IsOptionalPrompt.ToString());
					WriteAttributeString(writer, "Kind", pf.Kind.ToString());
					//WriteAttributeString(writer,"MaximumValue", (string) pf.MaximumValue);
					//WriteAttributeString(writer,"MinimumValue", (string) pf.MinimumValue);
					WriteAttributeString(writer, "Name", pf.Name);
					WriteAttributeString(writer, "NumberOfBytes", pf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
					WriteAttributeString(writer, "ParameterFieldName", pf.ParameterFieldName);
					WriteAttributeString(writer, "ParameterFieldUsage", pf.ParameterFieldUsage2.ToString());
					WriteAttributeString(writer, "ParameterType", pf.ParameterType.ToString());
					WriteAttributeString(writer, "ParameterValueKind", pf.ParameterValueKind.ToString());
					WriteAttributeString(writer, "PromptText", pf.PromptText);
					WriteAttributeString(writer, "ReportName", pf.ReportName);
					WriteAttributeString(writer, "ValueType", pf.ValueType.ToString());

					writer.WriteStartElement("ParameterDefaultValues");
					if (pf.DefaultValues.Count > 0)
					{
						foreach (ParameterValue pv in pf.DefaultValues)
						{
							writer.WriteStartElement("ParameterDefaultValue");
							WriteAttributeString(writer, "Description", pv.Description);
							// TODO: document dynamic parameters
							if (!pv.IsRange)
							{
								ParameterDiscreteValue pdv = (ParameterDiscreteValue)pv;
								WriteAttributeString(writer, "Value", pdv.Value.ToString());
							}
							writer.WriteEndElement();
						}
					}
					writer.WriteEndElement();

					writer.WriteStartElement("ParameterInitialValues");
					if (ddm_pf != null)
					{
						if (ddm_pf.InitialValues.Count > 0)
						{
							foreach (CRDataDefModel.ParameterFieldValue pv in ddm_pf.InitialValues)
							{
								writer.WriteStartElement("ParameterInitialValue");
								CRDataDefModel.ParameterFieldDiscreteValue pdv = (CRDataDefModel.ParameterFieldDiscreteValue)pv;
								WriteAttributeString(writer, "Value", pdv.Value.ToString());
								writer.WriteEndElement();
							}
						}
					}
					writer.WriteEndElement();

					writer.WriteStartElement("ParameterCurrentValues");
					if (pf.CurrentValues.Count > 0)
					{
						foreach (ParameterValue pv in pf.CurrentValues)
						{
							writer.WriteStartElement("ParameterCurrentValue");
							WriteAttributeString(writer, "Description", pv.Description);
							// TODO: document dynamic parameters
							if (!pv.IsRange)
							{
								ParameterDiscreteValue pdv = (ParameterDiscreteValue)pv;
								WriteAttributeString(writer, "Value", pdv.Value.ToString());
							}
							writer.WriteEndElement();
						}
					}
					writer.WriteEndElement();
				}

			}
			else if (fo is RunningTotalFieldDefinition rtf)
			{
                writer.WriteStartElement("RunningTotalFieldDefinition");
				//WriteAttributeString(writer,"EvaluationConditionType", rtf.EvaluationCondition);
				WriteAttributeString(writer, "EvaluationConditionType", rtf.EvaluationConditionType.ToString());
				WriteAttributeString(writer, "FormulaName", rtf.FormulaName);
				if (rtf.Group != null)
                {
                    WriteAttributeString(writer, "Group", rtf.Group.ToString());
                }

                WriteAttributeString(writer, "Kind", rtf.Kind.ToString());
				WriteAttributeString(writer, "Name", rtf.Name);
				WriteAttributeString(writer, "NumberOfBytes", rtf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "Operation", rtf.Operation.ToString());
				WriteAttributeString(writer, "OperationParameter", rtf.OperationParameter.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "ResetConditionType", rtf.ResetConditionType.ToString());

				if (rtf.SecondarySummarizedField != null)
                {
                    WriteAttributeString(writer, "SecondarySummarizedField", rtf.SecondarySummarizedField.FormulaName);
                }

                WriteAttributeString(writer, "SummarizedField", rtf.SummarizedField.FormulaName);
				WriteAttributeString(writer, "ValueType", rtf.ValueType.ToString());

			}
			else if (fo is SpecialVarFieldDefinition svf)
			{
				writer.WriteStartElement("SpecialVarFieldDefinition");
                WriteAttributeString(writer, "FormulaName", svf.FormulaName);
				WriteAttributeString(writer, "Kind", svf.Kind.ToString());
				WriteAttributeString(writer, "Name", svf.Name);
				WriteAttributeString(writer, "NumberOfBytes", svf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "SpecialVarType", svf.SpecialVarType.ToString());
				WriteAttributeString(writer, "ValueType", svf.ValueType.ToString());

			}
			else if (fo is SQLExpressionFieldDefinition sef)
			{
				writer.WriteStartElement("SQLExpressionFieldDefinition");

                WriteAttributeString(writer, "FormulaName", sef.FormulaName);
				WriteAttributeString(writer, "Kind", sef.Kind.ToString());
				WriteAttributeString(writer, "Name", sef.Name);
				WriteAttributeString(writer, "NumberOfBytes", sef.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "Text", sef.Text);
				WriteAttributeString(writer, "ValueType", sef.ValueType.ToString());

			}
			else if (fo is SummaryFieldDefinition sf)
			{
				writer.WriteStartElement("SummaryFieldDefinition");

                WriteAttributeString(writer, "FormulaName", sf.FormulaName);

				if (sf.Group != null)
                {
                    WriteAttributeString(writer, "Group", sf.Group.ToString());
                }

                WriteAttributeString(writer, "Kind", sf.Kind.ToString());
				WriteAttributeString(writer, "Name", sf.Name);
				WriteAttributeString(writer, "NumberOfBytes", sf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "Operation", sf.Operation.ToString());
				WriteAttributeString(writer, "OperationParameter", sf.OperationParameter.ToString(CultureInfo.InvariantCulture));
				if (sf.SecondarySummarizedField != null)
                {
                    WriteAttributeString(writer, "SecondarySummarizedField", sf.SecondarySummarizedField.ToString());
                }

                WriteAttributeString(writer, "SummarizedField", sf.SummarizedField.ToString());
				WriteAttributeString(writer, "ValueType", sf.ValueType.ToString());

			}
			writer.WriteEndElement();
		}

		private CRDataDefModel.ParameterField GetRASDDMParameterFieldObject(string fieldName, ReportDocument report)
		{
			CRDataDefModel.ParameterField rdm;
			if (report.IsSubreport)
			{
				var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
				rdm = subrptClientDoc.DataDefController.DataDefinition.ParameterFields.FindField(fieldName,
					CRDataDefModel.CrFieldDisplayNameTypeEnum.crFieldDisplayNameName) as CRDataDefModel.ParameterField;
			}
			else
			{
				rdm = _rcd.DataDefController.DataDefinition.ParameterFields.FindField(fieldName,
					CRDataDefModel.CrFieldDisplayNameTypeEnum.crFieldDisplayNameName) as CRDataDefModel.ParameterField;
			}
			return rdm;
		}

		private CRDataDefModel.FormulaField GetRASDDMFormulaFieldObject(string fieldName, ReportDocument report)
		{
			CRDataDefModel.FormulaField rdm;
			if (report.IsSubreport)
			{
				var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
				rdm = subrptClientDoc.DataDefController.DataDefinition.FormulaFields.FindField(fieldName,
					CRDataDefModel.CrFieldDisplayNameTypeEnum.crFieldDisplayNameName) as CRDataDefModel.FormulaField;
			}
			else
			{
				rdm = _rcd.DataDefController.DataDefinition.FormulaFields.FindField(fieldName,
					CRDataDefModel.CrFieldDisplayNameTypeEnum.crFieldDisplayNameName) as CRDataDefModel.FormulaField;
			}
			return rdm;
		}

		private void GetAreaFormat(Area area, XmlWriter writer)
		{
			writer.WriteStartElement("AreaFormat");

			WriteAttributeString(writer, "EnableHideForDrillDown", area.AreaFormat.EnableHideForDrillDown.ToString());
			WriteAttributeString(writer, "EnableKeepTogether", area.AreaFormat.EnableKeepTogether.ToString());
			WriteAttributeString(writer, "EnableNewPageAfter", area.AreaFormat.EnableNewPageAfter.ToString());
			WriteAttributeString(writer, "EnableNewPageBefore", area.AreaFormat.EnableNewPageBefore.ToString());
			WriteAttributeString(writer, "EnablePrintAtBottomOfPage", area.AreaFormat.EnablePrintAtBottomOfPage.ToString());
			WriteAttributeString(writer, "EnableResetPageNumberAfter", area.AreaFormat.EnableResetPageNumberAfter.ToString());
			WriteAttributeString(writer, "EnableSuppress", area.AreaFormat.EnableSuppress.ToString());

			if (area.Kind == AreaSectionKind.GroupHeader)
			{
				GroupAreaFormat gaf = (GroupAreaFormat)area.AreaFormat;
				writer.WriteStartElement("GroupAreaFormat");
				WriteAttributeString(writer, "EnableKeepGroupTogether", gaf.EnableKeepGroupTogether.ToString());
				WriteAttributeString(writer, "EnableRepeatGroupHeader", gaf.EnableRepeatGroupHeader.ToString());
				WriteAttributeString(writer, "VisibleGroupNumberPerPage", gaf.VisibleGroupNumberPerPage.ToString());
				writer.WriteEndElement();
			}
			writer.WriteEndElement();

		}

		private void GetBorderFormat(ReportObject ro, ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("Border");

			var border = ro.Border;
			WriteAttributeString(writer, "BottomLineStyle", border.BottomLineStyle.ToString());
			WriteAttributeString(writer, "HasDropShadow", border.HasDropShadow.ToString());
			WriteAttributeString(writer, "LeftLineStyle", border.LeftLineStyle.ToString());
			WriteAttributeString(writer, "RightLineStyle", border.RightLineStyle.ToString());
			WriteAttributeString(writer, "TopLineStyle", border.TopLineStyle.ToString());

			CRReportDefModel.ISCRReportObject rdmRo = GetRASRDMReportObject(ro.Name, report);
			if (rdmRo != null)
            {
                GetBorderConditionFormulas(rdmRo, writer);
            }

            if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
            {
                GetColorFormat(border.BackgroundColor, writer, "BackgroundColor");
            }

            if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
            {
                GetColorFormat(border.BorderColor, writer, "BorderColor");
            }

            writer.WriteEndElement();
		}

		private static void GetColorFormat(Color color, XmlWriter writer, String elementName = "Color")
		{
			writer.WriteStartElement(elementName);

			WriteAttributeString(writer, "Name", color.Name);
			WriteAttributeString(writer, "A", color.A.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "R", color.R.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "G", color.G.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "B", color.B.ToString(CultureInfo.InvariantCulture));

			writer.WriteEndElement();
		}

		private void GetFontFormat(Font font, XmlWriter writer)
		{
			writer.WriteStartElement("Font");

			WriteAttributeString(writer, "Bold", font.Bold.ToString());
			WriteAttributeString(writer, "FontFamily", font.FontFamily.Name);
			WriteAttributeString(writer, "GdiCharSet", font.GdiCharSet.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "GdiVerticalFont", font.GdiVerticalFont.ToString());
			WriteAttributeString(writer, "Height", font.Height.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "IsSystemFont", font.IsSystemFont.ToString());
			WriteAttributeString(writer, "Italic", font.Italic.ToString());
			WriteAttributeString(writer, "Name", font.Name);
			WriteAttributeString(writer, "OriginalFontName", font.OriginalFontName);
			WriteAttributeString(writer, "Size", font.Size.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "SizeinPoints", font.SizeInPoints.ToString(CultureInfo.InvariantCulture));
			WriteAttributeString(writer, "Strikeout", font.Strikeout.ToString());
			WriteAttributeString(writer, "Style", font.Style.ToString());
			WriteAttributeString(writer, "SystemFontName", font.SystemFontName);
			WriteAttributeString(writer, "Underline", font.Underline.ToString());
			WriteAttributeString(writer, "Unit", font.Unit.ToString());

			writer.WriteEndElement();
		}

		private void GetObjectFormat(ReportObject ro, XmlWriter writer)
		{
			writer.WriteStartElement("ObjectFormat");


			WriteAttributeString(writer, "CssClass", ro.ObjectFormat.CssClass);
			WriteAttributeString(writer, "EnableCanGrow", ro.ObjectFormat.EnableCanGrow.ToString());
			WriteAttributeString(writer, "EnableCloseAtPageBreak", ro.ObjectFormat.EnableCloseAtPageBreak.ToString());
			WriteAttributeString(writer, "EnableKeepTogether", ro.ObjectFormat.EnableKeepTogether.ToString());
			WriteAttributeString(writer, "EnableSuppress", ro.ObjectFormat.EnableSuppress.ToString());
			WriteAttributeString(writer, "HorizontalAlignment", ro.ObjectFormat.HorizontalAlignment.ToString());



			writer.WriteEndElement();
		}

		private void GetSectionFormat(Section section, ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("SectionFormat");

			WriteAttributeString(writer, "CssClass", section.SectionFormat.CssClass);
			WriteAttributeString(writer, "EnableKeepTogether", section.SectionFormat.EnableKeepTogether.ToString());
			WriteAttributeString(writer, "EnableNewPageAfter", section.SectionFormat.EnableNewPageAfter.ToString());
			WriteAttributeString(writer, "EnableNewPageBefore", section.SectionFormat.EnableNewPageBefore.ToString());
			WriteAttributeString(writer, "EnablePrintAtBottomOfPage", section.SectionFormat.EnablePrintAtBottomOfPage.ToString());
			WriteAttributeString(writer, "EnableResetPageNumberAfter", section.SectionFormat.EnableResetPageNumberAfter.ToString());
			WriteAttributeString(writer, "EnableSuppress", section.SectionFormat.EnableSuppress.ToString());
			WriteAttributeString(writer, "EnableSuppressIfBlank", section.SectionFormat.EnableSuppressIfBlank.ToString());
			WriteAttributeString(writer, "EnableUnderlaySection", section.SectionFormat.EnableUnderlaySection.ToString());

			CRReportDefModel.Section rdm_ro = GetRASRDMSectionObjectFromCRENGSectionObject(section.Name, report);
			if (rdm_ro != null)
            {
                GetSectionAreaFormatConditionFormulas(rdm_ro, writer);
            }


            if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
            {
                GetColorFormat(section.SectionFormat.BackgroundColor, writer, "BackgroundColor");
            }

            writer.WriteEndElement();
		}

		private void GetReportDefinition(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("ReportDefinition");

			GetAreas(report, writer);

			writer.WriteEndElement();
		}

		private void GetAreas(ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("Areas");

			foreach (Area area in report.ReportDefinition.Areas)
			{
				writer.WriteStartElement("Area");

				WriteAttributeString(writer, "Kind", area.Kind.ToString());
				WriteAttributeString(writer, "Name", area.Name);

				if ((ShowFormatTypes & FormatTypes.AreaFormat) == FormatTypes.AreaFormat)
                {
                    GetAreaFormat(area, writer);
                }

                GetSections(area, report, writer);

				writer.WriteEndElement();
			}

			writer.WriteEndElement();
		}

		private void GetSections(Area area, ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("Sections");

			foreach (Section section in area.Sections)
			{
				writer.WriteStartElement("Section");

				WriteAttributeString(writer, "Height", section.Height.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "Kind", section.Kind.ToString());
				WriteAttributeString(writer, "Name", section.Name);

				if ((ShowFormatTypes & FormatTypes.SectionFormat) == FormatTypes.SectionFormat)
                {
                    GetSectionFormat(section, report, writer);
                }

                GetReportObjects(section, report, writer);

				writer.WriteEndElement();
			}

			writer.WriteEndElement();
		}

		private void GetReportObjects(Section section, ReportDocument report, XmlWriter writer)
		{
			writer.WriteStartElement("ReportObjects");

			foreach (ReportObject reportObject in section.ReportObjects)
			{
				writer.WriteStartElement(reportObject.GetType().Name);

				CRReportDefModel.ISCRReportObject rasrdm_ro = GetRASRDMReportObject(reportObject.Name, report);

				WriteAttributeString(writer, "Name", reportObject.Name);
				WriteAttributeString(writer, "Kind", reportObject.Kind.ToString());

				WriteAttributeString(writer, "Top", reportObject.Top.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "Left", reportObject.Left.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "Width", reportObject.Width.ToString(CultureInfo.InvariantCulture));
				WriteAttributeString(writer, "Height", reportObject.Height.ToString(CultureInfo.InvariantCulture));
				if (reportObject is SubreportObject srobj)
				{
                    WriteAttributeString(writer, "SubreportName", srobj.SubreportName);
					WriteAttributeString(writer, "EnableOnDemand", srobj.EnableOnDemand.ToString(CultureInfo.InvariantCulture));

				}
				else if (reportObject is BoxObject bo)
				{
                    WriteAttributeString(writer, "Bottom", bo.Bottom.ToString(CultureInfo.InvariantCulture));
					WriteAttributeString(writer, "EnableExtendToBottomOfSection", bo.EnableExtendToBottomOfSection.ToString());
					WriteAttributeString(writer, "EndSectionName", bo.EndSectionName);
					WriteAttributeString(writer, "LineStyle", bo.LineStyle.ToString());
					WriteAttributeString(writer, "LineThickness", bo.LineThickness.ToString(CultureInfo.InvariantCulture));
					WriteAttributeString(writer, "Right", bo.Right.ToString(CultureInfo.InvariantCulture));
					if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
                    {
                        GetColorFormat(bo.LineColor, writer, "LineColor");
                    }
                }
				else if (reportObject is DrawingObject dobj)
				{
                    WriteAttributeString(writer, "Bottom", dobj.Bottom.ToString(CultureInfo.InvariantCulture));
					WriteAttributeString(writer, "EnableExtendToBottomOfSection", dobj.EnableExtendToBottomOfSection.ToString());
					WriteAttributeString(writer, "EndSectionName", dobj.EndSectionName);
					WriteAttributeString(writer, "LineStyle", dobj.LineStyle.ToString());
					WriteAttributeString(writer, "LineThickness", dobj.LineThickness.ToString(CultureInfo.InvariantCulture));
					WriteAttributeString(writer, "Right", dobj.Right.ToString(CultureInfo.InvariantCulture));
					if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
                    {
                        GetColorFormat(dobj.LineColor, writer, "LineColor");
                    }
                }
				else if (reportObject is FieldHeadingObject fh)
				{
                    var rasrdmFh = (CRReportDefModel.FieldHeadingObject)rasrdm_ro;
					WriteAttributeString(writer, "FieldObjectName", fh.FieldObjectName);
					WriteAttributeString(writer, "MaxNumberOfLines", rasrdmFh.MaxNumberOfLines.ToString());
					writer.WriteElementString("Text", fh.Text);

					if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
                    {
                        GetColorFormat(fh.Color, writer);
                    }

                    if ((ShowFormatTypes & FormatTypes.Font) == FormatTypes.Font)
					{
						GetFontFormat(fh.Font, writer);
						GetFontColorConditionFormulas(rasrdmFh.FontColor, writer);
					}
				}
				else if (reportObject is FieldObject fo)
				{
                    var rasrdmFo = (CRReportDefModel.FieldObject)rasrdm_ro;

					if (fo.DataSource != null)
                    {
                        WriteAttributeString(writer, "DataSource", fo.DataSource.FormulaName);
                    }

                    if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
                    {
                        GetColorFormat(fo.Color, writer);
                    }

                    if ((ShowFormatTypes & FormatTypes.Font) == FormatTypes.Font)
					{
						GetFontFormat(fo.Font, writer);
						GetFontColorConditionFormulas(rasrdmFo.FontColor, writer);
					}

				}
				else if (reportObject is TextObject tobj)
				{
                    var rasrdmTobj = (CRReportDefModel.TextObject)rasrdm_ro;

					WriteAttributeString(writer, "MaxNumberOfLines", rasrdmTobj.MaxNumberOfLines.ToString());
					writer.WriteElementString("Text", tobj.Text);

					if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
                    {
                        GetColorFormat(tobj.Color, writer);
                    }

                    if ((ShowFormatTypes & FormatTypes.Font) == FormatTypes.Font)
					{
						GetFontFormat(tobj.Font, writer);
						GetFontColorConditionFormulas(rasrdmTobj.FontColor, writer);
					}
				}

				if ((ShowFormatTypes & FormatTypes.Border) == FormatTypes.Border)
                {
                    GetBorderFormat(reportObject, report, writer);
                }

                if ((ShowFormatTypes & FormatTypes.ObjectFormat) == FormatTypes.ObjectFormat)
                {
                    GetObjectFormat(reportObject, writer);
                }


                if (rasrdm_ro != null)
                {
                    GetObjectFormatConditionFormulas(rasrdm_ro, writer);
                }

                writer.WriteEndElement();
			}

			writer.WriteEndElement();
		}

        public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (_report != null && _createdReport)
                {
                    _report.Dispose();
                }

                _report = null;
				_rcd = null;

				if (_oleCompoundFile != null)
				{
					((IDisposable)_oleCompoundFile).Dispose();
					_oleCompoundFile = null;
				}
			}
		}
	}
}
