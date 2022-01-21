# ExcelToPdf
excel export to pdf using dotnet core npoi

## How to use

1. using `NpoiExcelHelper` export excel to .html
2. using `PdfHelper` convert .html to pdf

## Using in asp.net core

> you must use `services.AddHtmlToPdf()` to inject, dot not new Converter() in your method.

1. in StartUp.cs

	```
	services.AddHtmlToPdf();
	```

2. in controller

	```

	private readonly IConverter _pdfConverter;

	/// <summary>
	/// ctor
	/// </summary>
	public TestController(IConverter pdfConverter)
	{
		_pdfConverter = pdfConverter;
	}

	/// <summary>
	/// ctor
	/// </summary>
	[HttpPost("Test")]
	public void TestMethod()
	{
		...
		_pdfConverter.HtmlToPdf(htmlFilePath, pdfFilePath);
		...
	}
	```

## in docker

to be continued...

## sample

to be continued...