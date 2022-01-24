# ExcelToPdf

[中文版](/README.md) | [English](./doc/README_En.md)

This is a simple project to export Excel to PDF.

It consists of two steps:
1. Use ` NpoiExcelHelper. ExcelToHtml ` method will excel export to HTML
2. Use `Pdfhelper.HtmlToPdf` to convert HTML to PDF

Npoi and WkHtmlToPdf did almost all the work, and because of that, please go to these two projects to view the relevant documents, thank them very much

- [WkHtmlToPdf-DotNet](https://github.com/HakanL/WkHtmlToPdf-DotNet)
- [NPOI](https://github.com/nissl-lab/npoi/wiki/How-to-use-NPOI-on-Linux)

This project also provides a simple encapsulation of customizations, see the customization section of the documentation.

## Sample

For details, please refer to `ExcelToPdfSample` project.

```csharp
public void OnPostSample2()
{
	var excelFileInfo = new FileInfo("TestData/sample.xls");
	var htmlFileInfo = new FileInfo("Output/sample.html");
	var pdfFileInfo = new FileInfo("Output/sample.pdf");

	if (htmlFileInfo.Directory != null && !htmlFileInfo.Directory.Exists)
	{
		htmlFileInfo.Directory.Create();
	}

	// export excel to html
	NpoiExcelHelper.ExcelToHtml(excelFileInfo.FullName, htmlFileInfo.FullName, configOptions: option =>
	{
		option.OutputColumnHeaders = true;
	});

	// convert html to pdf
	_converter.HtmlToPdf(htmlFileInfo.FullName, pdfFileInfo.FullName,
		config =>
		{
			config.Orientation = Orientation.Landscape;
		});
}
```

> Note that using asp.net requires the injection of `IConverter` first, as follows:
```csharp
in StartUp.cs

// injection
services.AddHtmlToPdf();
```

## Custom

### Custom excel export

With the `NpoiExcelHelper.ExcelToHtml` method, can use the configOptions parameters to HTML do custom processing for export, specific definition refer to:[Npoi ExcelToHtmlConverter](https://github.com/nissl-lab/npoi/blob/edac37ddf7c442e8e66b47f72d53d9aa81c5db35/ooxml/SS/Converter/ExcelToHtmlConverter.cs)


### Custom pdf convert

With the `PdfHelper.HtmlToPdf` method, can use configGlobalSettings parameter to custom pdf:
```csharp
// ... other code ...
_converter.HtmlToPdf(htmlFileInfo.FullName, pdfFileInfo.FullName,
	config =>
	{
		config.Orientation = Orientation.Portrait;
	});
// ... other code ...
```

**GlobalSettings Commonly used attributes：**

|parameter|comment|default|note|
|-|-|-|-|
|Orientation|he orientation of the output document|portrait|Portrait, Landscape|
|ColorMode|Should the output be printed in color or gray scale|color||
|UseCompression|Should we use loss less compression when creating the pdf file.|true||
|DPI|What dpi should we use when printing|96|
|DocumentTitle|The title of the PDF document|empty||
|ImageDPI|The maximal DPI to use for images in the pdf document|600||
|imageQuality|The jpeg compression factor to use when producing the pdf document|94||
|PaperSize|Size of output paper||
|PaperWidth|The height of the output document||
|PaperHeight|The width of the output document||
|MarginLeft|Size of the left margin||
|MarginRight|Size of the right margin||
|MarginTop|Size of the top margin||
|MarginBottom|Size of the bottom margin||

## Docker and linux

[How-to-use-NPOI-on-Linux](https://github.com/nissl-lab/npoi/wiki/How-to-use-NPOI-on-Linux)
