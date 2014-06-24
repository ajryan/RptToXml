using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Xml;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportAppServer.ClientDoc;
using CrystalDecisions.ReportAppServer.Controllers;
using CrystalDecisions.Shared;
using CRDataDefModel = CrystalDecisions.ReportAppServer.DataDefModel;
using CRReportDefModel = CrystalDecisions.ReportAppServer.ReportDefModel;

//TODO: avoid Double Typecast antipattern.  http://www.boyet.com/Articles/DoubleCastingAntiPattern.html
//Maybe use this syntax http://stackoverflow.com/a/7252992/270794

namespace RptToXml
{
	public partial class RptDefinitionWriter: IDisposable
	{
		private const FormatTypes ShowFormatTypes = FormatTypes.AreaFormat | FormatTypes.SectionFormat | FormatTypes.Color;

		private ReportDocument _report;
		private ISCDReportClientDocument _rcd;
		private bool _createdReport;

		public RptDefinitionWriter(string filename)
		{
			_createdReport = true;
			_report = new ReportDocument();
			_report.Load(filename, OpenReportMethod.OpenReportByTempCopy);
			_rcd = _report.ReportClientDocument;

			Trace.WriteLine("Loaded report");
		}

		public RptDefinitionWriter(ReportDocument value)
		{
			_report = value;
		}

		public void WriteToXml(System.IO.Stream output)
		{
			using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings { Indent = true }))
			{
				WriteToXml(writer);
			}
		}

		public void WriteToXml(string targetXmlPath)
		{
			using (XmlWriter writer = XmlWriter.Create(targetXmlPath, new XmlWriterSettings { Indent = true }))
			{
				WriteToXml(writer);
			}
		}

		public void WriteToXml(XmlWriter writer)
		{
			Trace.WriteLine("Writing to XML");

			writer.WriteStartDocument();
			ProcessReport(_report, writer);
			writer.WriteEndDocument();
			writer.Flush();
		}

		//This is a recursive method.  GetSubreports() calls it.
		private void ProcessReport(ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Report");

			writer.WriteAttributeString("Name", report.Name);
			Trace.WriteLine("Writing report " + report.Name);


			if (!report.IsSubreport)
			{
				Trace.WriteLine("Writing header info");

				writer.WriteAttributeString("FileName", report.FileName.Replace("rassdk://", ""));
				writer.WriteAttributeString("HasSavedData", report.HasSavedData.ToString());

				GetSummaryinfo(report, writer);
				GetReportOptions(report, writer);
				GetPrintOptions(report, writer);
				GetSubreports(report, writer);  //recursion happens here.
			}

			GetDatabase(report, writer);
			GetDataDefinition(report, writer);
			GetReportDefinition(report, writer);

			writer.WriteEndElement();
		}

		private static void GetSummaryinfo(ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Summaryinfo");

			writer.WriteAttributeString("KeywordsinReport", report.SummaryInfo.KeywordsInReport);
			writer.WriteAttributeString("ReportAuthor", report.SummaryInfo.ReportAuthor);
			writer.WriteAttributeString("ReportComments", report.SummaryInfo.ReportComments);
			writer.WriteAttributeString("ReportSubject", report.SummaryInfo.ReportSubject);
			writer.WriteAttributeString("ReportTitle", report.SummaryInfo.ReportTitle);

			writer.WriteEndElement();
		}

		private static void GetReportOptions(ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "ReportOptions");

			writer.WriteAttributeString("EnableSaveDataWithReport", report.ReportOptions.EnableSaveDataWithReport.ToString());
			writer.WriteAttributeString("EnableSavePreviewPicture", report.ReportOptions.EnableSavePreviewPicture.ToString());
			writer.WriteAttributeString("EnableSaveSummariesWithReport", report.ReportOptions.EnableSaveSummariesWithReport.ToString());
			writer.WriteAttributeString("EnableUseDummyData", report.ReportOptions.EnableUseDummyData.ToString());
			writer.WriteAttributeString("initialDataContext", report.ReportOptions.InitialDataContext);
			writer.WriteAttributeString("initialReportPartName", report.ReportOptions.InitialDataContext);

			writer.WriteEndElement();
		}

		private void GetPrintOptions(ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "PrintOptions");

			writer.WriteAttributeString("PageContentHeight", report.PrintOptions.PageContentHeight.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("PageContentWidth", report.PrintOptions.PageContentWidth.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("PaperOrientation", report.PrintOptions.PaperOrientation.ToString());
			writer.WriteAttributeString("PaperSize", report.PrintOptions.PaperSize.ToString());
			writer.WriteAttributeString("PaperSource", report.PrintOptions.PaperSource.ToString());
			writer.WriteAttributeString("PrinterDuplex", report.PrintOptions.PrinterDuplex.ToString());
			writer.WriteAttributeString("PrinterName", report.PrintOptions.PrinterName);

			WriteAndTraceStartElement(writer, "PageMargins");

			writer.WriteAttributeString("bottomMargin", report.PrintOptions.PageMargins.bottomMargin.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("leftMargin", report.PrintOptions.PageMargins.leftMargin.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("rightMargin", report.PrintOptions.PageMargins.rightMargin.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("topMargin", report.PrintOptions.PageMargins.topMargin.ToString(CultureInfo.InvariantCulture));

			writer.WriteEndElement();

			CRReportDefModel.PrintOptions rdmPrintOptions = GetRASRDMPrintOptionsObject(report.Name, report);
			if (rdmPrintOptions != null)
				GetPageMarginConditionFormulas(rdmPrintOptions, writer);

			writer.WriteEndElement();
		}

		private void GetSubreports(ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "SubReports");

			foreach (ReportDocument subreport in report.Subreports)
				ProcessReport(subreport, writer);

			writer.WriteEndElement();
		}

		private void GetDatabase(ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Database");

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
			WriteAndTraceStartElement(writer, "TableLinks");

			foreach (TableLink tl in report.Database.Links)
			{
				WriteAndTraceStartElement(writer, "TableLink");
				writer.WriteAttributeString("JoinType", tl.JoinType.ToString());

				WriteAndTraceStartElement(writer, "SourceFields");
				foreach (FieldDefinition fd in tl.SourceFields)
					GetFieldDefinition(fd, writer);
				writer.WriteEndElement();

				WriteAndTraceStartElement(writer, "DestinationFields");
				foreach (FieldDefinition fd in tl.DestinationFields)
					GetFieldDefinition(fd, writer);
				writer.WriteEndElement();

				writer.WriteEndElement();
			}

			writer.WriteEndElement();
		}

		private void GetReportClientTables(ISCDReportClientDocument reportClientDocument, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Tables");
			
			foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table table in reportClientDocument.DatabaseController.Database.Tables)
			{
				GetTable(table, writer);
			}

			writer.WriteEndElement();
		}
		private void GetSubreportClientTables(SubreportClientDocument subrptClientDocument, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Tables");

			foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table table in subrptClientDocument.DatabaseController.Database.Tables)
			{
				GetTable(table, writer);
			}

			writer.WriteEndElement();
		}

		private void GetTable(CrystalDecisions.ReportAppServer.DataDefModel.Table table, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Table");

			writer.WriteAttributeString("Alias", table.Alias);
			writer.WriteAttributeString("ClassName", table.ClassName);
			writer.WriteAttributeString("Name", table.Name);

			WriteAndTraceStartElement(writer, "ConnectionInfo");
			foreach (string propertyId in table.ConnectionInfo.Attributes.PropertyIDs)
			{
				// make attribute name safe for XML
				string attributeName = propertyId.Replace(" ", "_");

				writer.WriteAttributeString(attributeName, table.ConnectionInfo.Attributes[propertyId].ToString());
			}
			
			writer.WriteAttributeString("UserName", table.ConnectionInfo.UserName);
			writer.WriteAttributeString("Password", table.ConnectionInfo.Password);
			writer.WriteEndElement();

			var commandTable = table as CRDataDefModel.CommandTable;
			if (commandTable != null)
			{
				var cmdTable = commandTable;
				WriteAndTraceStartElement(writer, "Command");
				writer.WriteString(cmdTable.CommandText);
				writer.WriteEndElement();
			}

			WriteAndTraceStartElement(writer, "Fields");

			foreach (CrystalDecisions.ReportAppServer.DataDefModel.Field fd in table.DataFields)
			{
				GetFieldDefinition(fd, writer);
			}

			writer.WriteEndElement();

			writer.WriteEndElement();
		}

		private static void GetFieldDefinition(FieldDefinition fd, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Field");

			writer.WriteAttributeString("FormulaName", fd.FormulaName);
			writer.WriteAttributeString("Kind", fd.Kind.ToString());
			writer.WriteAttributeString("Name", fd.Name);
			writer.WriteAttributeString("NumberOfBytes", fd.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("ValueType", fd.ValueType.ToString());

			writer.WriteEndElement();
		}

		private static void GetFieldDefinition(CrystalDecisions.ReportAppServer.DataDefModel.Field fd, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Field");

			writer.WriteAttributeString("Description", fd.Description);
			writer.WriteAttributeString("FormulaForm", fd.FormulaForm);
			writer.WriteAttributeString("HeadingText", fd.HeadingText);
			writer.WriteAttributeString("IsRecurring", fd.IsRecurring.ToString());
			writer.WriteAttributeString("Kind", fd.Kind.ToString());
			writer.WriteAttributeString("Length", fd.Length.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("LongName", fd.LongName);
			writer.WriteAttributeString("Name", fd.Name);
			writer.WriteAttributeString("ShortName", fd.ShortName);
			writer.WriteAttributeString("Type", fd.Type.ToString());
			writer.WriteAttributeString("UseCount", fd.UseCount.ToString(CultureInfo.InvariantCulture));
			
			writer.WriteEndElement();
		}

		private void GetDataDefinition(ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "DataDefinition");

			writer.WriteElementString("GroupSelectionFormula", report.DataDefinition.GroupSelectionFormula);
			writer.WriteElementString("RecordSelectionFormula", report.DataDefinition.RecordSelectionFormula);

			WriteAndTraceStartElement(writer, "Groups");
			foreach (Group group in report.DataDefinition.Groups)
			{
				WriteAndTraceStartElement(writer, "Group");
				writer.WriteAttributeString("ConditionField", group.ConditionField.FormulaName);

				//TODO: Not sure how to properly reference the GroupOptions from DDM.  romanows
				//CRDataDefModel.GroupOptions rdm_go = GetRASDDMGroupOptionsObject(report);
				//if (rdm_ro != null)
				//    GetGroupOptionsConditionFormulas(rdm_go, writer);
				
				writer.WriteEndElement();

			}
			writer.WriteEndElement();

			WriteAndTraceStartElement(writer, "SortFields");
			foreach (SortField sortField in report.DataDefinition.SortFields)
			{
				WriteAndTraceStartElement(writer, "SortField");

				writer.WriteAttributeString("Field", sortField.Field.FormulaName);
				try
				{
					string sortDirection = sortField.SortDirection.ToString();
					writer.WriteAttributeString("SortDirection", sortDirection);
				}
				catch (NotSupportedException)
				{}
				writer.WriteAttributeString("SortType", sortField.SortType.ToString());

				writer.WriteEndElement();
			}
			writer.WriteEndElement();

			WriteAndTraceStartElement(writer, "FormulaFieldDefinitions");
			foreach (var field in report.DataDefinition.FormulaFields)
				GetFieldObject(field, writer);
			writer.WriteEndElement();

			WriteAndTraceStartElement(writer, "GroupNameFieldDefinitions");
			foreach (var field in report.DataDefinition.GroupNameFields)
				GetFieldObject(field, writer);
			writer.WriteEndElement();

			WriteAndTraceStartElement(writer, "ParameterFieldDefinitions");
			foreach (var field in report.DataDefinition.ParameterFields)
				GetFieldObject(field, writer);
			writer.WriteEndElement();

			WriteAndTraceStartElement(writer, "RunningTotalFieldDefinitions");
			foreach (var field in report.DataDefinition.RunningTotalFields)
				GetFieldObject(field, writer);
			writer.WriteEndElement();

			WriteAndTraceStartElement(writer, "SQLExpressionFields");
			foreach (var field in report.DataDefinition.SQLExpressionFields)
				GetFieldObject(field, writer);
			writer.WriteEndElement();

			WriteAndTraceStartElement(writer, "SummaryFields");
			foreach (var field in report.DataDefinition.SummaryFields)
				GetFieldObject(field, writer);
			writer.WriteEndElement();

			writer.WriteEndElement();
		}

		private void GetFieldObject(Object fo, XmlWriter writer)
		{
			if (fo is DatabaseFieldDefinition)
			{
				var df = (DatabaseFieldDefinition)fo;

				WriteAndTraceStartElement(writer, "DatabaseFieldDefinition");

				writer.WriteAttributeString("FormulaName", df.FormulaName);
				writer.WriteAttributeString("Kind", df.Kind.ToString());
				writer.WriteAttributeString("Name", df.Name);
				writer.WriteAttributeString("NumberOfBytes", df.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("TableName", df.TableName);
				writer.WriteAttributeString("ValueType", df.ValueType.ToString());

				writer.WriteEndElement();
			}
			else if (fo is FormulaFieldDefinition)
			{
				var ff = (FormulaFieldDefinition)fo;

				WriteAndTraceStartElement(writer, "FormulaFieldDefinition");

				writer.WriteAttributeString("FormulaName", ff.FormulaName);
				writer.WriteAttributeString("Kind", ff.Kind.ToString());
				writer.WriteAttributeString("Name", ff.Name);
				writer.WriteAttributeString("NumberOfBytes", ff.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("ValueType", ff.ValueType.ToString());
				writer.WriteString(ff.Text);

				writer.WriteEndElement();
			}
			else if (fo is GroupNameFieldDefinition)
			{
				var gnf = (GroupNameFieldDefinition)fo;

				WriteAndTraceStartElement(writer, "GroupNameFieldDefinition");

				writer.WriteAttributeString("FormulaName", gnf.FormulaName);
				writer.WriteAttributeString("Group", gnf.Group.ToString());
				writer.WriteAttributeString("GroupNameFieldName", gnf.GroupNameFieldName);
				writer.WriteAttributeString("Kind", gnf.Kind.ToString());
				writer.WriteAttributeString("Name", gnf.Name);
				writer.WriteAttributeString("NumberOfBytes", gnf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("ValueType", gnf.ValueType.ToString());

				writer.WriteEndElement();
			}
			else if (fo is ParameterFieldDefinition)
			{
				var pf = (ParameterFieldDefinition)fo;

				WriteAndTraceStartElement(writer, "ParameterFieldDefinition");

				writer.WriteAttributeString("EditMask", pf.EditMask);
				writer.WriteAttributeString("EnableAllowEditingDefaultValue", pf.EnableAllowEditingDefaultValue.ToString());
				writer.WriteAttributeString("EnableAllowMultipleValue", pf.EnableAllowMultipleValue.ToString());
				writer.WriteAttributeString("EnableNullValue", pf.EnableNullValue.ToString());
				writer.WriteAttributeString("FormulaName", pf.FormulaName);
				writer.WriteAttributeString("HasCurrentValue", pf.HasCurrentValue.ToString());
				writer.WriteAttributeString("Kind", pf.Kind.ToString());
				//writer.WriteAttributeString("MaximumValue", (string) pf.MaximumValue);
				//writer.WriteAttributeString("MinimumValue", (string) pf.MinimumValue);
				writer.WriteAttributeString("Name", pf.Name);
				writer.WriteAttributeString("NumberOfBytes", pf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("ParameterFieldName", pf.ParameterFieldName);
				writer.WriteAttributeString("ParameterFieldUsage", pf.ParameterFieldUsage2.ToString());
				writer.WriteAttributeString("ParameterType", pf.ParameterType.ToString());
				writer.WriteAttributeString("ParameterValueKind", pf.ParameterValueKind.ToString());
				writer.WriteAttributeString("PromptText", pf.PromptText);
				writer.WriteAttributeString("ReportName", pf.ReportName);
				writer.WriteAttributeString("ValueType", pf.ValueType.ToString());

				writer.WriteEndElement();
			}
			else if (fo is RunningTotalFieldDefinition)
			{
				var rtf = (RunningTotalFieldDefinition)fo;

				WriteAndTraceStartElement(writer, "RunningTotalFieldDefinition");
				//writer.WriteAttributeString("EvaluationConditionType", rtf.EvaluationCondition);
				writer.WriteAttributeString("EvaluationConditionType", rtf.EvaluationConditionType.ToString());
				writer.WriteAttributeString("FormulaName", rtf.FormulaName);
				if (rtf.Group != null) writer.WriteAttributeString("Group", rtf.Group.ToString());
				writer.WriteAttributeString("Kind", rtf.Kind.ToString());
				writer.WriteAttributeString("Name", rtf.Name);
				writer.WriteAttributeString("NumberOfBytes", rtf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("Operation", rtf.Operation.ToString());
				writer.WriteAttributeString("OperationParameter", rtf.OperationParameter.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("ResetConditionType", rtf.ResetConditionType.ToString());
				
				if (rtf.SecondarySummarizedField != null)
					writer.WriteAttributeString("SecondarySummarizedField", rtf.SecondarySummarizedField.FormulaName);

				writer.WriteAttributeString("SummarizedField", rtf.SummarizedField.FormulaName);
				writer.WriteAttributeString("ValueType", rtf.ValueType.ToString());

				writer.WriteEndElement();
			}
			else if (fo is SpecialVarFieldDefinition)
			{
				WriteAndTraceStartElement(writer, "SpecialVarFieldDefinition");
				var svf = (SpecialVarFieldDefinition)fo;
				writer.WriteAttributeString("FormulaName", svf.FormulaName);
				writer.WriteAttributeString("Kind", svf.Kind.ToString());
				writer.WriteAttributeString("Name", svf.Name);
				writer.WriteAttributeString("NumberOfBytes", svf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("SpecialVarType", svf.SpecialVarType.ToString());
				writer.WriteAttributeString("ValueType", svf.ValueType.ToString());

				writer.WriteEndElement();
			}
			else if (fo is SQLExpressionFieldDefinition)
			{
				WriteAndTraceStartElement(writer, "SQLExpressionFieldDefinition");
				var sef = (SQLExpressionFieldDefinition)fo;

				writer.WriteAttributeString("FormulaName", sef.FormulaName);
				writer.WriteAttributeString("Kind", sef.Kind.ToString());
				writer.WriteAttributeString("Name", sef.Name);
				writer.WriteAttributeString("NumberOfBytes", sef.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("Text", sef.Text);
				writer.WriteAttributeString("ValueType", sef.ValueType.ToString());

				writer.WriteEndElement();
			}
			else if (fo is SummaryFieldDefinition)
			{
				WriteAndTraceStartElement(writer, "SummaryFieldDefinition");

				var sf = (SummaryFieldDefinition)fo;

				writer.WriteAttributeString("FormulaName", sf.FormulaName);
				
				if (sf.Group != null)
					writer.WriteAttributeString("Group", sf.Group.ToString());

				writer.WriteAttributeString("Kind", sf.Kind.ToString());
				writer.WriteAttributeString("Name", sf.Name);
				writer.WriteAttributeString("NumberOfBytes", sf.NumberOfBytes.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("Operation", sf.Operation.ToString());
				writer.WriteAttributeString("OperationParameter", sf.OperationParameter.ToString(CultureInfo.InvariantCulture));
				if (sf.SecondarySummarizedField != null) writer.WriteAttributeString("SecondarySummarizedField", sf.SecondarySummarizedField.ToString());
				writer.WriteAttributeString("SummarizedField", sf.SummarizedField.ToString());
				writer.WriteAttributeString("ValueType", sf.ValueType.ToString());

				writer.WriteEndElement();
			}
		}

		private void GetAreaFormat(Area area, ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "AreaFormat");

			writer.WriteAttributeString("EnableHideForDrillDown", area.AreaFormat.EnableHideForDrillDown.ToString());
			writer.WriteAttributeString("EnableKeepTogether", area.AreaFormat.EnableKeepTogether.ToString());
			writer.WriteAttributeString("EnableNewPageAfter", area.AreaFormat.EnableNewPageAfter.ToString());
			writer.WriteAttributeString("EnableNewPageBefore", area.AreaFormat.EnableNewPageBefore.ToString());
			writer.WriteAttributeString("EnablePrintAtBottomOfPage", area.AreaFormat.EnablePrintAtBottomOfPage.ToString());
			writer.WriteAttributeString("EnableResetPageNumberAfter", area.AreaFormat.EnableResetPageNumberAfter.ToString());
			writer.WriteAttributeString("EnableSuppress", area.AreaFormat.EnableSuppress.ToString());

			writer.WriteEndElement();
		}

		private void GetBorderFormat(ReportObject ro, ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Border");

			var border = ro.Border;
			writer.WriteAttributeString("BottomLineStyle", border.BottomLineStyle.ToString());
			writer.WriteAttributeString("HasDropShadow", border.HasDropShadow.ToString());
			writer.WriteAttributeString("LeftLineStyle", border.LeftLineStyle.ToString());
			writer.WriteAttributeString("RightLineStyle", border.RightLineStyle.ToString());
			writer.WriteAttributeString("TopLineStyle", border.TopLineStyle.ToString());

			CRReportDefModel.ISCRReportObject rdm_ro = GetRASRDMReportObjectFromCRENGReportObject(ro.Name, report);
			if (rdm_ro != null)
				GetBorderConditionFormulas(rdm_ro, writer);
			
			if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
				GetColorFormat(border.BackgroundColor, writer, "BackgroundColor");
			if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
				GetColorFormat(border.BorderColor, writer, "BorderColor");

			writer.WriteEndElement();
		}

		private static void GetColorFormat(Color color, XmlWriter writer, String elementName = "Color")
		{
			WriteAndTraceStartElement(writer, elementName);

			writer.WriteAttributeString("Name", color.Name);
			writer.WriteAttributeString("A", color.A.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("R", color.R.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("G", color.G.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("B", color.B.ToString(CultureInfo.InvariantCulture));

			writer.WriteEndElement();
		}

		private void GetFontFormat(Font font, ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Font");

			writer.WriteAttributeString("Bold", font.Bold.ToString());
			writer.WriteAttributeString("FontFamily", font.FontFamily.Name);
			writer.WriteAttributeString("GdiCharSet", font.GdiCharSet.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("GdiVerticalFont", font.GdiVerticalFont.ToString());
			writer.WriteAttributeString("Height", font.Height.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("IsSystemFont", font.IsSystemFont.ToString());
			writer.WriteAttributeString("Italic", font.Italic.ToString());
			writer.WriteAttributeString("Name", font.Name);
			writer.WriteAttributeString("OriginalFontName", font.OriginalFontName);
			writer.WriteAttributeString("Size", font.Size.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("SizeinPoints", font.SizeInPoints.ToString(CultureInfo.InvariantCulture));
			writer.WriteAttributeString("Strikeout", font.Strikeout.ToString());
			writer.WriteAttributeString("Style", font.Style.ToString());
			writer.WriteAttributeString("SystemFontName", font.SystemFontName);
			writer.WriteAttributeString("Underline", font.Underline.ToString());
			writer.WriteAttributeString("Unit", font.Unit.ToString());
			if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
				GetFontColorConditionFormulas();

			writer.WriteEndElement();
		}

		private void GetObjectFormat(ObjectFormat objectFormat, ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "ObjectFormat");

			writer.WriteAttributeString("CssClass", objectFormat.CssClass);
			writer.WriteAttributeString("EnableCanGrow", objectFormat.EnableCanGrow.ToString());
			writer.WriteAttributeString("EnableCloseAtPageBreak", objectFormat.EnableCloseAtPageBreak.ToString());
			writer.WriteAttributeString("EnableKeepTogether", objectFormat.EnableKeepTogether.ToString());
			writer.WriteAttributeString("EnableSuppress", objectFormat.EnableSuppress.ToString());
			writer.WriteAttributeString("HorizontalAlignment", objectFormat.HorizontalAlignment.ToString());

			GetObjectFormatConditionFormulas();

			writer.WriteEndElement();
		}

		private void GetSectionFormat(Section section, ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "SectionFormat");

			writer.WriteAttributeString("CssClass", section.SectionFormat.CssClass);
			writer.WriteAttributeString("EnableKeepTogether", section.SectionFormat.EnableKeepTogether.ToString());
			writer.WriteAttributeString("EnableNewPageAfter", section.SectionFormat.EnableNewPageAfter.ToString());
			writer.WriteAttributeString("EnableNewPageBefore", section.SectionFormat.EnableNewPageBefore.ToString());
			writer.WriteAttributeString("EnablePrintAtBottomOfPage", section.SectionFormat.EnablePrintAtBottomOfPage.ToString());
			writer.WriteAttributeString("EnableResetPageNumberAfter", section.SectionFormat.EnableResetPageNumberAfter.ToString());
			writer.WriteAttributeString("EnableSuppress", section.SectionFormat.EnableSuppress.ToString());
			writer.WriteAttributeString("EnableSuppressIfBlank", section.SectionFormat.EnableSuppressIfBlank.ToString());
			writer.WriteAttributeString("EnableUnderlaySection", section.SectionFormat.EnableUnderlaySection.ToString());

			CRReportDefModel.Section rdm_ro = GetRASRDMSectionObjectFromCRENGSectionObject(section.Name, report);
			if (rdm_ro != null)
				GetSectionAreaFormatConditionFormulas(rdm_ro, writer);
		

			if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
				GetColorFormat(section.SectionFormat.BackgroundColor, writer, "BackgroundColor");

			writer.WriteEndElement();
		}

		private void GetReportDefinition(ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "ReportDefinition");

			GetAreas(report, writer);

			writer.WriteEndElement();
		}

		private void GetAreas(ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Areas");

			foreach (Area area in report.ReportDefinition.Areas)
			{
				WriteAndTraceStartElement(writer, "Area");

				writer.WriteAttributeString("Kind", area.Kind.ToString());
				writer.WriteAttributeString("Name", area.Name);

				if ((ShowFormatTypes & FormatTypes.AreaFormat) == FormatTypes.AreaFormat)
					GetAreaFormat(area, report, writer);
				GetSections(area, report, writer);

				writer.WriteEndElement();
			}

			writer.WriteEndElement();
		}

		private void GetSections(Area area, ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "Sections");

			foreach (Section section in area.Sections)
			{
				WriteAndTraceStartElement(writer, "Section");


				writer.WriteAttributeString("Height", section.Height.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("Kind", section.Kind.ToString());
				writer.WriteAttributeString("Name", section.Name);

				if ((ShowFormatTypes & FormatTypes.SectionFormat) == FormatTypes.SectionFormat)
					GetSectionFormat(section, report, writer);

				GetReportObjects(section, report, writer);

				writer.WriteEndElement();
			}

			writer.WriteEndElement();
		}

		private void GetReportObjects(Section section, ReportDocument report, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "ReportObjects");

			foreach (ReportObject reportObject in section.ReportObjects)
			{
				WriteAndTraceStartElement(writer, reportObject.GetType().Name);

				writer.WriteAttributeString("Name", reportObject.Name);
				writer.WriteAttributeString("Kind", reportObject.Kind.ToString());

				writer.WriteAttributeString("Top", reportObject.Top.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("Left", reportObject.Left.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("Width", reportObject.Width.ToString(CultureInfo.InvariantCulture));
				writer.WriteAttributeString("Height", reportObject.Height.ToString(CultureInfo.InvariantCulture));

				if (reportObject is BoxObject)
				{
					var bo = (BoxObject)reportObject;
					writer.WriteAttributeString("Bottom", bo.Bottom.ToString(CultureInfo.InvariantCulture));
					writer.WriteAttributeString("EnableExtendToBottomOfSection", bo.EnableExtendToBottomOfSection.ToString());
					writer.WriteAttributeString("EndSectionName", bo.EndSectionName);
					writer.WriteAttributeString("LineStyle", bo.LineStyle.ToString());
					writer.WriteAttributeString("LineThickness", bo.LineThickness.ToString(CultureInfo.InvariantCulture));
					writer.WriteAttributeString("Right", bo.Right.ToString(CultureInfo.InvariantCulture));
					if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
						GetColorFormat(bo.LineColor, writer, "LineColor");
				}
				else if (reportObject is DrawingObject)
				{
					var dobj = (DrawingObject)reportObject;
					writer.WriteAttributeString("Bottom", dobj.Bottom.ToString(CultureInfo.InvariantCulture));
					writer.WriteAttributeString("EnableExtendToBottomOfSection", dobj.EnableExtendToBottomOfSection.ToString());
					writer.WriteAttributeString("EndSectionName", dobj.EndSectionName);
					writer.WriteAttributeString("LineStyle", dobj.LineStyle.ToString());
					writer.WriteAttributeString("LineThickness", dobj.LineThickness.ToString(CultureInfo.InvariantCulture));
					writer.WriteAttributeString("Right", dobj.Right.ToString(CultureInfo.InvariantCulture));
					if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
						GetColorFormat(dobj.LineColor, writer, "LineColor");
				}
				else if (reportObject is FieldHeadingObject)
				{
					var fh = (FieldHeadingObject)reportObject;
					writer.WriteAttributeString("FieldObjectName", fh.FieldObjectName);
					writer.WriteElementString("Text", fh.Text);
				}
				else if (reportObject is FieldObject)
				{
					var fo = (FieldObject)reportObject;

					if (fo.DataSource != null)
						writer.WriteAttributeString("DataSource", fo.DataSource.FormulaName);

					if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
						GetColorFormat(fo.Color, writer);

					if ((ShowFormatTypes & FormatTypes.Font) == FormatTypes.Font)
						GetFontFormat(fo.Font, report, writer);

				}
				else if (reportObject is LineObject)
				{
					var lo = (LineObject)reportObject;
					writer.WriteAttributeString("Bottom", lo.Bottom.ToString(CultureInfo.InvariantCulture));
					writer.WriteAttributeString("EnableExtendToBottomOfSection", lo.EnableExtendToBottomOfSection.ToString());
					writer.WriteAttributeString("EndSectionName", lo.EndSectionName);
					writer.WriteAttributeString("LineStyle", lo.LineStyle.ToString());
					writer.WriteAttributeString("LineThickness", lo.LineThickness.ToString(CultureInfo.InvariantCulture));
					writer.WriteAttributeString("Right", lo.Right.ToString(CultureInfo.InvariantCulture));
					if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
						GetColorFormat(lo.LineColor, writer, "LineColor");
				}
				else if (reportObject is TextObject)
				{
					var tobj = (TextObject)reportObject;
					writer.WriteElementString("Text", tobj.Text);

					if ((ShowFormatTypes & FormatTypes.Color) == FormatTypes.Color)
						GetColorFormat(tobj.Color, writer);
					if ((ShowFormatTypes & FormatTypes.Font) == FormatTypes.Font)
						GetFontFormat(tobj.Font, report, writer);
				}

				if ((ShowFormatTypes & FormatTypes.Border) == FormatTypes.Border)
					GetBorderFormat(reportObject, report, writer);

				if ((ShowFormatTypes & FormatTypes.ObjectFormat) == FormatTypes.ObjectFormat)
					GetObjectFormat(reportObject.ObjectFormat, report, writer);

				writer.WriteEndElement();
			}

			writer.WriteEndElement();
		}

		// pretty much straight from api docs
		private CommonFieldFormat GetCommonFieldFormat(string reportObjectName, ReportDocument report)
		{
			FieldObject field = report.ReportDefinition.ReportObjects[reportObjectName] as FieldObject;
			if (field != null)
			{
				return field.FieldFormat.CommonFormat;
			}
			return null;
		}

		private static void WriteAndTraceStartElement(XmlWriter writer, string elementName)
		{
			Trace.WriteLine("  " + elementName);
			writer.WriteStartElement(elementName);
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
					_report.Close();
					_report.Dispose();
					_report = null;
				}
			}
		}

		~RptDefinitionWriter()
		{
			Dispose(false);
		}
	}
}
