using System;
using System.Xml;

using CrystalDecisions.CrystalReports.Engine;

using CRDataDefModel = CrystalDecisions.ReportAppServer.DataDefModel;
using CRReportDefModel = CrystalDecisions.ReportAppServer.ReportDefModel;

namespace RptToXml
{
	public partial class RptDefinitionWriter
	{
		#region Get ReportAppServer Objects

		private CRReportDefModel.ISCRReportObject GetRASRDMReportObject(string oname, ReportDocument report)
		{
			CRReportDefModel.ISCRReportObject reportObj;
			if (report.IsSubreport)
			{
				var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
				reportObj = subrptClientDoc.ReportDefController.ReportDefinition.FindObjectByName(oname);
			}
			else
			{
				reportObj = _rcd.ReportDefController.ReportDefinition.FindObjectByName(oname);
			}
			return reportObj;
		}

		private CRReportDefModel.Section GetRASRDMSectionObjectFromCRENGSectionObject(string sname, ReportDocument report)
		{
			CRReportDefModel.Section section;
			if (report.IsSubreport)
			{
				var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
				section = subrptClientDoc.ReportDefController.ReportDefinition.FindSectionByName(sname);
			}
			else
			{
				section = _rcd.ReportDefController.ReportDefinition.FindSectionByName(sname);
			}
			return section;
		}

		private GroupAreaFormat GetRASDDMGroupAreaFormatObject(Group group, ReportDocument report)
		{
			//TODO:  finish me, not sure how to reference GroupOptions object.  The GroupOptionsConditionFormulas
			// has 2 forms, one is sort and group and the other is the collection that we are used to seeing.  It is confusing.
			GroupAreaFormat gaf = null;

			if (report.IsSubreport)
			{
				//var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
				//var groups = subrptClientDoc.DataDefController.GroupController.FindGroup(fieldName);
				//gaf = subrptClientDoc.ReportDefController.ReportDefinition.FindObjectByName();
			}
			else
			{
				//gaf = _rcd;
			}

			return gaf;
		}

		private CRReportDefModel.PrintOptions GetRASRDMPrintOptionsObject(string name, ReportDocument report)
		{
			if (report.IsSubreport)
				return null;

			return _rcd.ReportDocument.PrintOptions;
		}
	
		#endregion Get ReportAppServer Objects

		private static void GetBorderConditionFormulas(CRReportDefModel.ISCRReportObject ro, XmlWriter writer)
		{
			writer.WriteStartElement("BorderConditionFormulas");
			var cfs = Enum.GetValues(typeof(CRReportDefModel.CrBorderConditionFormulaTypeEnum));
			foreach (CRReportDefModel.CrBorderConditionFormulaTypeEnum cf in cfs)
			{
				var formula = ro.Border.ConditionFormulas[cf];

				if (!String.IsNullOrEmpty(formula.Text))
					writer.WriteAttributeString(GetShortEnumName(cf), formula.Text);
			}
			writer.WriteEndElement();
		}

		private void GetPageMarginConditionFormulas(CRReportDefModel.PrintOptions po, XmlWriter writer)
		{
			writer.WriteStartElement("PageMarginConditionFormulas");
			var cfs = Enum.GetValues(typeof(CRReportDefModel.CrPageMarginConditionFormulaTypeEnum));
			foreach (CRReportDefModel.CrPageMarginConditionFormulaTypeEnum cf in cfs)
			{
				var formula = po.PageMargins.PageMarginConditionFormulas[cf];
				if (!String.IsNullOrEmpty(formula.Text))
					writer.WriteAttributeString(GetShortEnumName(cf), formula.Text);
			}
			writer.WriteEndElement();
		}

		private static void GetSectionAreaFormatConditionFormulas(CRReportDefModel.Section ro, XmlWriter writer)
		{
			writer.WriteStartElement("SectionAreaConditionFormulas");
			var cfs = Enum.GetValues(typeof(CRReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum));

			//TODO: need to cut this down by Area.Kind to only show relevant attributes, i.e. Page Clamp is only valid on Page Footer.
			foreach (CRReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum cf in cfs)
			{
				var formula = ro.Format.ConditionFormulas[cf];
				if (!String.IsNullOrEmpty(formula.Text))
					writer.WriteAttributeString(GetShortEnumName(cf, "crSectionAreaConditionFormulaType"), formula.Text);
			}
			writer.WriteEndElement();
		}

		private static void GetGroupOptionsConditionFormulas(GroupOptions rdm_ro, XmlWriter writer)
		{
			//TODO: not yet implemented
		}

		private static void GetFontColorConditionFormulas(CRReportDefModel.FontColor fco, XmlWriter writer)
		{
			writer.WriteStartElement("FontColorConditionFormulas");

			foreach (var fontColorTypeObj in Enum.GetValues(typeof(CRReportDefModel.CrFontColorConditionFormulaTypeEnum)))
			{
				var fontColorType = (CRReportDefModel.CrFontColorConditionFormulaTypeEnum)fontColorTypeObj;

				var cf = fco.ConditionFormulas[fontColorType];

				if (!String.IsNullOrEmpty(cf.Text))
					writer.WriteAttributeString(GetShortEnumName(fontColorType), cf.Text);
			}

			writer.WriteEndElement();
		}

		private static void GetObjectFormatConditionFormulas(CRReportDefModel.ISCRReportObject ro, XmlWriter writer)
		{
			writer.WriteStartElement("ObjectFormatConditionFormulas");

			foreach (var formulaTypeObj in Enum.GetValues(typeof(CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum)))
			{
				var formulaType = (CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum)formulaTypeObj;

				var cf = ro.Format.ConditionFormulas[formulaType];

				if (!String.IsNullOrEmpty(cf.Text))
					writer.WriteAttributeString(GetShortEnumName(formulaType), cf.Text);
			}

            if (ro is CRReportDefModel.PictureObject)
            {
                var ro_p = (CRReportDefModel.PictureObject)ro;
                var cf = ro_p.GraphicLocationFormula;

                if (!String.IsNullOrEmpty(cf.Text))
                    writer.WriteAttributeString("GraphicLocation", cf.Text);
            }

			writer.WriteEndElement();
		}

		private static void GetTopNSortClassConditionFormulas()
		{
			//TODO: not yet implemented
		}

		private static string GetShortEnumName<T>(T enumValue, string prefix = null) where T : struct
		{
			if (prefix == null)
			{
				string typeName = typeof(T).Name;
				prefix = typeName.EndsWith("Enum", StringComparison.OrdinalIgnoreCase)
					? typeName.Substring(0, typeName.Length - 4)
					: typeName;
			}

			string valueString = enumValue.ToString();

			if (valueString.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
				valueString = valueString.Substring(prefix.Length);

			return valueString;
		}
	}
}
