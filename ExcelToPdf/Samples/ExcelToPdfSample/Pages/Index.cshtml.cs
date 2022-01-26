using ExcelToPdf;
using ExcelToPdf.Helper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using WkHtmlToPdfDotNet;
using WkHtmlToPdfDotNet.Contracts;

namespace ExcelToPdfSample.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly ExcelToPdfService _excelToPdfService;

        public IndexModel(ILogger<IndexModel> logger, ExcelToPdfService excelToPdfService)
        {
            _logger = logger;
            _excelToPdfService = excelToPdfService;
        }

        public void OnGet()
        {

        }

        public string Message { get; set; }

        public void OnPostSimpleExport()
        {
            FileInfo excelFileInfo, htmlFileInfo, pdfFileInfo;
            SetupSample(out excelFileInfo, out htmlFileInfo, out pdfFileInfo);

            // export pdf
            _excelToPdfService.ExportToPdf(excelFileInfo.FullName, pdfFileInfo.FullName);

            this.Message = $"Export Success, Pdf file save to: {pdfFileInfo.FullName}";
        }

        public void OnPostCustomPdf()
        {
            FileInfo excelFileInfo, htmlFileInfo, pdfFileInfo;
            SetupSample(out excelFileInfo, out htmlFileInfo, out pdfFileInfo);

            // export pdf
            _excelToPdfService.ExportToPdf(excelFileInfo.FullName, pdfFileInfo.FullName, configPdfGlobalSettings: config =>
            {
                config.Orientation = Orientation.Landscape;
                config.ColorMode = ColorMode.Grayscale;
                config.Margins.Left = 150;
                config.Margins.Top = 50;
            });

            this.Message = $"Export Success, Pdf file save to: {pdfFileInfo.FullName}";
        }

        /// <summary>
        /// 2 Step Export
        /// </summary>
        public void OnPostExport2Setp()
        {
            FileInfo excelFileInfo, htmlFileInfo, pdfFileInfo;
            SetupSample(out excelFileInfo, out htmlFileInfo, out pdfFileInfo);

            // 1. export excel to html
            NpoiExcelHelper.ExcelToHtml(excelFileInfo.FullName, htmlFileInfo.FullName);

            // 2. convert html to pdf
            _excelToPdfService.PdfConvert.HtmlToPdf(htmlFileInfo.FullName, pdfFileInfo.FullName,
                config =>
                {
                    config.Orientation = Orientation.Portrait;
                });

            this.Message = $"Export Success, Pdf file save to: {pdfFileInfo.FullName}";
        }

        /// <summary>
        /// prepare file path
        /// </summary>
        /// <param name="excelFileInfo"></param>
        /// <param name="htmlFileInfo"></param>
        /// <param name="pdfFileInfo"></param>
        private static void SetupSample(out FileInfo excelFileInfo, out FileInfo htmlFileInfo, out FileInfo pdfFileInfo)
        {
            excelFileInfo = new FileInfo("TestData/sample.xlsx");
            htmlFileInfo = new FileInfo("Output/sample.html");
            pdfFileInfo = new FileInfo("Output/sample.pdf");
            if (!htmlFileInfo.Directory.Exists)
            {
                htmlFileInfo.Directory.Create();
            }
        }
    }
}