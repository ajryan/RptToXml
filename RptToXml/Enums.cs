using System;

namespace RptToXml
{
	[Flags]
	public enum FormatTypes
	{
		None = 0,
		Border = 2 ^ 0,
		Color = 2 ^ 1,
		Font = 2 ^ 2,
		AreaFormat = 2 ^ 3,
		FieldFormat = 2 ^ 4,
		ObjectFormat = 2 ^ 5,
		SectionFormat = 2 ^ 6,
		All = Border & Color & Font & AreaFormat & FieldFormat & ObjectFormat & SectionFormat
	}

	[Flags]
	public enum ObjectTypes
	{
		None = 0,
		Area = 2 ^ 0,
		Section = 2 ^ 1,
		ReportObject = 2 ^ 2,
		All = Area & Section & ReportObject
	}
}