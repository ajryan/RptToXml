# RptToXml #

Dumps a Crystal Reports RPT file to XML. Useful for diffs.

Binary releases available on the [Releases](https://github.com/ajryan/RptToXml/releases) page.

Ported to C# from the [original VB project](http://code.google.com/p/rpttoxml/)

## Building ##

RptToXml references Crystal Reports assemblies. The easiest way to get them onto a development machine is to install the Crystal Reports Runtime from an MSI downloaded from [this page](http://scn.sap.com/docs/DOC-7824). The most recent support pack of Crystal Reports 13 should work.

The solution should build with VS2012 or higher. Express editions have not been tested but should work.
