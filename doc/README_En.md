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

## Install

```
Install-Package ExcelToPdf
```

## Sample

For details, please refer to `ExcelToPdfSample` project.

### Dependency Injection
```csharp
// in StartUp.cs
services.AddExcelToPdf("temp");
```

### Export

```csharp
private readonly ILogger<IndexModel> _logger;
private readonly ExcelToPdfService _excelToPdfService;

// inject ExcelToPdfService from controller
public IndexModel(ILogger<IndexModel> logger, ExcelToPdfService excelToPdfService)
{
	_logger = logger;
	_excelToPdfService = excelToPdfService;
}

// export to pdf
public void OnPostSimpleExport()
{
	FileInfo excelFileInfo, htmlFileInfo, pdfFileInfo;
	SetupSample(out excelFileInfo, out htmlFileInfo, out pdfFileInfo);

	// export pdf
	_excelToPdfService.ExportToPdf(excelFileInfo.FullName, pdfFileInfo.FullName);

	this.Message = $"Export Success, Pdf file save to: {pdfFileInfo.FullName}";
}
```

## Custom

### Custom excel export to html

With the `NpoiExcelHelper.ExcelToHtml` method, can use the configOptions parameters to HTML do custom processing for export, specific definition refer to:[Npoi ExcelToHtmlConverter](https://github.com/nissl-lab/npoi/blob/edac37ddf7c442e8e66b47f72d53d9aa81c5db35/ooxml/SS/Converter/ExcelToHtmlConverter.cs)


### Custom pdf convert

With the `ExportToPdf` method, can use configGlobalSettings parameter to custom pdf:
```csharp
// ... other code ...
// export pdf
_excelToPdfService.ExportToPdf(excelFileInfo.FullName, pdfFileInfo.FullName, configPdfGlobalSettings: config =>
{
	config.Orientation = Orientation.Landscape;
	config.ColorMode = ColorMode.Grayscale;
	config.Margins.Left = 150;
	config.Margins.Top = 50;
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
