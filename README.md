# Instruction

[中文版](./README.md) | [English](./doc/README_En.md)

这是一个简单的项目，用来将excel导出为pdf。

它由两个步骤组成：
1. 使用`NpoiExcelHelper.ExcelToHtml`方法将excel导出为html
2. 使用`PdfHelper.HtmlToPdf`转换html为pdf

npoi和WkHtmlToPdf做了几乎所有的工作，也因为这样，相关的文档请移步到这两个项目查看，特别感谢
- [WkHtmlToPdf-DotNet](https://github.com/HakanL/WkHtmlToPdf-DotNet)
- [NPOI](https://github.com/nissl-lab/npoi/wiki/How-to-use-NPOI-on-Linux)

## 示例

具体请参考`ExcelToPdfSample/Pages/Index.cshtml.cs`。

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

> 注意在asp.net中使用需要先注入`IConverter`，如下:
```csharp
in StartUp.cs

// 注入
services.AddHtmlToPdf();
```

## 自定义

### 自定义excel导出

使用`NpoiExcelHelper.ExcelToHtml`方法时，可以用configOptions参数来对导出的html做自定义处理，具体定义请参考：[Npoi ExcelToHtmlConverter](https://github.com/nissl-lab/npoi/blob/edac37ddf7c442e8e66b47f72d53d9aa81c5db35/ooxml/SS/Converter/ExcelToHtmlConverter.cs)


### 自定义pdf转换

使用`PdfHelper.HtmlToPdf`方法时，可以使用configGlobalSettings参数来定义pdf导出，具体请参考：
[WkHtmlToPdf-DotNet](https://github.com/HakanL/WkHtmlToPdf-DotNet)

## Docker以及linux下使用

[How-to-use-NPOI-on-Linux](https://github.com/nissl-lab/npoi/wiki/How-to-use-NPOI-on-Linux)
