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
		private CRReportDefModel.ISCRReportObject GetRASRDMReportObject(string oname, ReportDocument report)
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

		private GroupAreaFormat GetRASDDMGroupAreaFormatObject(Group group, ReportDocument report)
		{
			//TODO:  finish me, not sure how to reference GroupOptions object.  The GroupOptionsConditionFormulas
			// has 2 forms, one is sort and group and the other is the collection that we are used to seeing.  It is confusing.
			GroupAreaFormat gaf = null;
			if (report.IsSubreport)
			{
				var subrptClientDoc = _report.ReportClientDocument.SubreportController.GetSubreport(report.Name);
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

			//TODO: need to cut this down by Area.Kind to only show relevant attributes, i.e. Page Clamp is only valid on Page Footer. 
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

		private static void GetObjectFormatConditionFormulas(CRReportDefModel.ISCRReportObject ro, XmlWriter writer)
		{
			WriteAndTraceStartElement(writer, "ObjectFormatConditionFormulas");

			var cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableSuppress];
			if (cf.Text != null)
				writer.WriteAttributeString("EnableSuppress", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableKeepTogether];
			if (cf.Text != null)
				writer.WriteAttributeString("EnableKeepTogether", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableCloseAtPageBreak];
			if (cf.Text != null)
				writer.WriteAttributeString("EnableCloseAtPageBreak", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeHorizontalAlignment];
			if (cf.Text != null)
				writer.WriteAttributeString("HorizontalAlignment", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableCanGrow];
			if (cf.Text != null)
				writer.WriteAttributeString("EnableCanGrow", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeToolTipText];
			if (cf.Text != null)
				writer.WriteAttributeString("ToolTipText", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeRotation];
			if (cf.Text != null)
				writer.WriteAttributeString("Rotation", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeHyperlink];
			if (cf.Text != null)
				writer.WriteAttributeString("Hyperlink", cf.Text.ToString());
			
			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableSuppress];
			if (cf.Text != null)
				writer.WriteAttributeString("CssClass", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDisplayString];
			if (cf.Text != null)
				writer.WriteAttributeString("DisplayString", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDeltaX];
			if (cf.Text != null)
				writer.WriteAttributeString("DeltaX", cf.Text.ToString());

			cf = ro.Format.ConditionFormulas[CRReportDefModel.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDeltaWidth];
			if (cf.Text != null)
				writer.WriteAttributeString("DeltaWidth", cf.Text.ToString());

			writer.WriteEndElement();
		}

		private static void GetTopNSortClassConditionFormulas()
		{
			//TODO: throw new NotImplementedException();
		}

	}
}
