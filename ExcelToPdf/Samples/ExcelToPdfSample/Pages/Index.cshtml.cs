using ExcelToPdf;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using WkHtmlToPdfDotNet;
using WkHtmlToPdfDotNet.Contracts;

namespace ExcelToPdfSample.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly IConverter _converter;

        public IndexModel(ILogger<IndexModel> logger, IConverter converter)
        {
            _logger = logger;
            _converter = converter;
        }

        public void OnGet()
        {

        }

        public void OnPostSample1()
        {
            var excelFileInfo = new FileInfo("TestData/sample.xlsx");
            var htmlFileInfo = new FileInfo("Output/sample.html");
            var pdfFileInfo = new FileInfo("Output/sample.pdf");

            if (!htmlFileInfo.Directory.Exists)
            {
                htmlFileInfo.Directory.Create();
            }

            NpoiExcelHelper.ExcelToHtml(excelFileInfo.FullName, htmlFileInfo.FullName);
            _converter.HtmlToPdf(htmlFileInfo.FullName, pdfFileInfo.FullName);
        }

        public void OnPostSample2()
        {
            var excelFileInfo = new FileInfo("TestData/sample2.xls");
            var htmlFileInfo = new FileInfo("Output/sample2.html");
            var pdfFileInfo = new FileInfo("Output/sample2.pdf");

            if (htmlFileInfo.Directory != null && !htmlFileInfo.Directory.Exists)
            {
                htmlFileInfo.Directory.Create();
            }

            NpoiExcelHelper.ExcelToHtml(excelFileInfo.FullName, htmlFileInfo.FullName);
            _converter.HtmlToPdf(htmlFileInfo.FullName, pdfFileInfo.FullName,
                config =>
                {
                    config.Orientation = Orientation.Landscape;
                });
        }
    }
}