using ExcelToPdf.Helper;
using NPOI.SS.Converter;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WkHtmlToPdfDotNet;
using WkHtmlToPdfDotNet.Contracts;

namespace ExcelToPdf
{
    public class ExcelToPdfService
    {
        private readonly IConverter _pdfConvert;
        private readonly string _tempDir;

        public ExcelToPdfService(IConverter pdfConvert, string tempDir)
        {
            _pdfConvert = pdfConvert;
            _tempDir = tempDir;
        }

        /// <summary>
        /// Pdf Converter
        /// </summary>
        public IConverter PdfConvert => _pdfConvert;

        /// <summary>
        /// export pdf
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="pdfFilePath"></param>
        /// <param name="removeSheetName"></param>
        /// <param name="configOptions"></param>
        /// <param name="afterProcess"></param>
        /// <param name="configPdfGlobalSettings"></param>
        public void ExportToPdf(IWorkbook workbook, string pdfFilePath, bool removeSheetName = true, Action<ExcelToHtmlConverter> configOptions = null, Action<ExcelToHtmlConverter> afterProcess = null, Action<GlobalSettings>? configPdfGlobalSettings = null)
        {
            // setup temp directory
            var tempDir = new DirectoryInfo(_tempDir);
            if (!tempDir.Exists)
            {
                tempDir.Create();
            }

            // 1. export html first
            var htmlFilePath = Path.Combine(tempDir.FullName, $"temp_excel_html_{Guid.NewGuid}.html");
            NpoiExcelHelper.ExcelToHtml(workbook, htmlFilePath, removeSheetName, configOptions, afterProcess);

            // 2. convert html to pdf
            PdfConvert.HtmlToPdf(htmlFilePath, pdfFilePath, configPdfGlobalSettings);

            // remove tmp html file
            File.Delete(htmlFilePath);
        }

        /// <summary>
        /// export pdf using excel file
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <param name="pdfFilePath"></param>
        /// <param name="removeSheetName"></param>
        /// <param name="configOptions"></param>
        /// <param name="afterProcess"></param>
        /// <param name="configPdfGlobalSettings"></param>
        public void ExportToPdf(string excelFilePath, string pdfFilePath, bool removeSheetName = true, Action<ExcelToHtmlConverter> configOptions = null, Action<ExcelToHtmlConverter> afterProcess = null, Action<GlobalSettings>? configPdfGlobalSettings = null)
        {
            var workbook = NpoiExcelHelper.GetWorkbook(excelFilePath);
            ExportToPdf(workbook, pdfFilePath, removeSheetName, configOptions, afterProcess, configPdfGlobalSettings);
        }
    }
}
