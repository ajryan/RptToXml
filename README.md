# RptToXml

Dumps a Crystal Reports RPT file to XML. Useful for diffs.

Binary releases available on the [Releases](https://github.com/ajryan/RptToXml/releases) page.

Ported to C# from the [original VB project](http://code.google.com/p/rpttoxml/)

## Running

Download the latest [release](https://github.com/ajryan/RptToXml/releases).

RptToXml references Crystal Reports assemblies. The easiest way to get them onto a development machine is to install the Crystal Reports Runtime from an MSI downloaded from [this page](https://www.sap.com/cmp/td/sap-crystal-reports-visual-studio-trial.html).

Install the SAP frameworks (probably need 32-bit and 64-bit):

- SAP Crystal Reports for Visual Studio (SP##) runtime engine for .NET framework MSI (32-bit)
- SAP Crystal Reports for Visual Studio (SP##) runtime engine for .NET framework MSI (64-bit)

Run the executable from the command line with

```sh
# process a single file
RptToXml.exe path/to/report_name.rpt path/to/output.xml

# process a directory of reports
cd path/to/reports
RptToXml.exe -r
```

## Building From Source

The solution will build with VS2012 or higher. Express editions have not been tested but should work.

Install the SAP pacakge for Visual Studio:

- SAP Crystal Reports for Visual Studio (SP##) installation package for Microsoft Visual Studio IDE (VS 20## and above or below)


Find the executable `RptToXml.exe` in ```RptToXml/bin/<where did you build to?>``` after building the solution in Visual Studio.
