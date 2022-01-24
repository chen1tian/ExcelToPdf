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

        public string Message { get; set; }

        public void OnPostSample()
        {
            var excelFileInfo = new FileInfo("TestData/sample.xlsx");
            var htmlFileInfo = new FileInfo("Output/sample.html");
            var pdfFileInfo = new FileInfo("Output/sample.pdf");

            if (!htmlFileInfo.Directory.Exists)
            {
                htmlFileInfo.Directory.Create();
            }

            // export excel to html
            NpoiExcelHelper.ExcelToHtml(excelFileInfo.FullName, htmlFileInfo.FullName);

            // convert html to pdf
            _converter.HtmlToPdf(htmlFileInfo.FullName, pdfFileInfo.FullName,
                config =>
                {
                    config.Orientation = Orientation.Landscape;
                });

            this.Message = $"Export Success, Pdf file save to: {pdfFileInfo.FullName}";
        }
    }
}