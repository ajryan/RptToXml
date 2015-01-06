using System;
using System.Xml;
using System.Text;
using CrystalDecisions.CrystalReports.Engine;
using CRDataDefModel = CrystalDecisions.ReportAppServer.DataDefModel;
using CRReportDefModel = CrystalDecisions.ReportAppServer.ReportDefModel;

namespace RptToXml
{
	public partial class RptDefinitionWriter : IDisposable
	{

		#region Get ReportAppServer Objects
		private CRReportDefModel.ISCRReportObject GetRASRDMReportObjectFromCRENGReportObject(string oname, ReportDocument report)
		{
			CRReportDefModel.ISCRReportObject rdm;
			if (report.IsSubreport)
			{
				var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
				rdm = subrptClientDoc.ReportDefController.ReportDefinition.FindObjectByName(oname) as CRReportDefModel.ISCRReportObject;
			}
			else
			{
				rdm = _rcd.ReportDefController.ReportDefinition.FindObjectByName(oname) as CRReportDefModel.ISCRReportObject;
			}
			return rdm;
		}

		private CRReportDefModel.Section GetRASRDMSectionObjectFromCRENGSectionObject(string sname, ReportDocument report)
		{
			CRReportDefModel.Section rdm;
			if (report.IsSubreport)
			{
				var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
				rdm = subrptClientDoc.ReportDefController.ReportDefinition.FindSectionByName(sname) as CRReportDefModel.Section;
			}
			else
			{
				rdm = _rcd.ReportDefController.ReportDefinition.FindSectionByName(sname) as CRReportDefModel.Section;
			}
			return rdm;
		}

		private CRDataDefModel.GroupOptions GetRASDDMGroupOptionsObject(ReportDocument report)
		{
			//TODO:  finish me, not sure how to reference GroupOptions object.  The GroupOptionsConditionFormulas
			// has 2 forms, one is sort and group and the other is the collection that we are used to seeing.  It is confusing.
			CRDataDefModel.GroupOptions rdm = null;
			if (report.IsSubreport)
			{
				//var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
			}
			else
			{}
				//rdm = _rcd.ReportDocument.GroupOptions as CRDataDefModel.GroupOptions;

			return rdm;
		}

		private CRReportDefModel.PrintOptions GetRASRDMPrintOptionsObject(string name, ReportDocument report)
		{
			CRReportDefModel.PrintOptions rdm;
			if (report.IsSubreport)
				throw new NotImplementedException();
			else
				rdm = _rcd.ReportDocument.PrintOptions as CRReportDefModel.PrintOptions;

			return rdm;
		}
		#endregion Get ReportAppServer Objects

		private static void GetBorderConditionFormulas(CRReportDefModel.ISCRReportObject ro, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "BorderConditionFormulas");
			var cfs = Enum.GetValues(typeof(CRReportDefModel.CrBorderConditionFormulaTypeEnum));
			foreach (CRReportDefModel.CrBorderConditionFormulaTypeEnum cf in cfs)
			{
				var formula = ro.Border.ConditionFormulas[cf];
				var text = (formula.Text == null) ? "" : formula.Text.ToString();
				var enumname = Enum.GetName(typeof(CRReportDefModel.CrBorderConditionFormulaTypeEnum), cf);
				var shortenumname = enumname.Substring("crBorderConditionFormulaType".Length);
				writer.WriteAttributeString(shortenumname, text);
			}
			writer.WriteEndElement();
		}

		private void GetPageMarginConditionFormulas(CRReportDefModel.PrintOptions po, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "PageMarginConditionFormulas");
			var cfs = Enum.GetValues(typeof(CRReportDefModel.CrPageMarginConditionFormulaTypeEnum));
			foreach (CRReportDefModel.CrPageMarginConditionFormulaTypeEnum cf in cfs)
			{
				var formula = po.PageMargins.PageMarginConditionFormulas[cf];
				var text = (formula.Text == null) ? "" : formula.Text.ToString();
				var enumname = Enum.GetName(typeof(CRReportDefModel.CrPageMarginConditionFormulaTypeEnum), cf);
				var shortenumname = enumname.Substring("CrPageMarginConditionFormulaType".Length);
				writer.WriteAttributeString(shortenumname, text);
			}
			writer.WriteEndElement();
		}

		private static void GetSectionAreaFormatConditionFormulas(CRReportDefModel.Section ro, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "SectionAreaConditionFormulas");
			var cfs = Enum.GetValues(typeof(CRReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum));

			//TODO: need to cut this down by Area.Kind to only shjow relevant attributes, i.e. Page Clamp is only valid on Page Footer. 
			foreach (CRReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum cf in cfs)
			{
				var formula = ro.Format.ConditionFormulas[cf];
				var text = (formula.Text == null) ? "" : formula.Text.ToString();
				var enumname = Enum.GetName(typeof(CRReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum), cf);
				var shortenumname = enumname.Substring("crSectionAreaConditionFormulaType".Length);
				writer.WriteAttributeString(shortenumname, text);
			}
			writer.WriteEndElement();
		}

		private static void GetGroupOptionsConditionFormulas(GroupOptions rdm_ro, XmlWriter writer)
		{
			//TODO: throw new NotImplementedException();
		}

		private static void GetFontColorConditionFormulas()
		{
			//TODO: throw new NotImplementedException();
		}

		private static void GetObjectFormatConditionFormulas()
		{
			//TODO: throw new NotImplementedException();
		}

		private static void GetTopNSortClassConditionFormulas()
		{
			//TODO: throw new NotImplementedException();
		}

	}
}
