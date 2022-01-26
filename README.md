# 说明

[中文版](/README.md) | [English](./doc/README_En.md)

这是一个简单的项目，用来将excel导出为pdf。

它由两个步骤组成：
1. 使用`NpoiExcelHelper.ExcelToHtml`方法将excel导出为html
2. 使用`PdfHelper.HtmlToPdf`转换html为pdf

npoi和WkHtmlToPdf做了几乎所有的工作，也因为这样，需要查询文档时请移步到这两个项目查看。
特别感谢
- [WkHtmlToPdf-DotNet](https://github.com/HakanL/WkHtmlToPdf-DotNet)
- [NPOI](https://github.com/nissl-lab/npoi/wiki/How-to-use-NPOI-on-Linux)

本项目也对自定义做了简单的封装，参考文档自定义一节。

## 安装

```
Install-Package ExcelToPdf -Version 6.0.0
```

## 示例

具体请参考`ExcelToPdfSample/Pages/Index.cshtml.cs`。

### 依赖注入
```csharp
// in StartUp.cs
services.AddExcelToPdf("temp");
```

### 导出
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

## 自定义

### 自定义excel导出为Html

使用`NpoiExcelHelper.ExcelToHtml`方法时，可以用configOptions参数来对导出的html做自定义处理，具体定义请参考：[Npoi ExcelToHtmlConverter](https://github.com/nissl-lab/npoi/blob/edac37ddf7c442e8e66b47f72d53d9aa81c5db35/ooxml/SS/Converter/ExcelToHtmlConverter.cs)


### 自定义pdf转换

使用`PdfHelper.HtmlToPdf`方法时，可以使用configGlobalSettings参数来定义pdf导出:
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

**GlobalSettings 常用属性：**

|参数名|说明|默认值|备注|
|-|-|-|-|
|Orientation|排版方向|竖向|Portrait：竖排，Landscape：横排|
|ColorMode|颜色模式|彩色|Color：彩色，Grayscale:灰度|
|UseCompression|使用压缩|是||
|DPI|打印dpi|96|
|DocumentTitle|文档标题|空||
|ImageDPI|图片DPI|600||
|imageQuality|图片质量|94||
|PaperSize|纸面尺寸||
|PaperWidth|文档宽度||
|PaperHeight|文档高度||
|MarginLeft|左边距||
|MarginRight|右边距||
|MarginTop|上边距||
|MarginBottom|下边距||

全部参数请参考：
[WkHtmlToPdf GlobalSettings源码](https://github.com/HakanL/WkHtmlToPdf-DotNet/blob/master/src/WkHtmlToPdf-DotNet/Settings/GlobalSettings.cs)

## Docker以及linux下使用

[How-to-use-NPOI-on-Linux](https://github.com/nissl-lab/npoi/wiki/How-to-use-NPOI-on-Linux)
