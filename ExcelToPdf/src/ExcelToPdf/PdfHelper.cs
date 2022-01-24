using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using SixLabors.Fonts;
using System;
using System.Globalization;
using System.IO;
using WkHtmlToPdfDotNet;
using WkHtmlToPdfDotNet.Contracts;

namespace ExcelToPdf
{
    /// <summary>
    /// Pdf帮助器
    /// </summary>
    public static class PdfHelper
    {
        /// <summary>
        /// Html转Pdf
        /// </summary>
        /// <param name="htmlFilePath"></param>
        /// <param name="pdfFilePath"></param>
        /// <param name="converter">转换器</param>
        /// <param name="configGlobalSettings"></param>
        public static void HtmlToPdf(this IConverter converter, string htmlFilePath, string pdfFilePath, Action<GlobalSettings>? configGlobalSettings = null)
        {
            // 默认配置
            var globalSettings = new GlobalSettings()
            {
                ColorMode = ColorMode.Color,
                Orientation = Orientation.Portrait,
                PaperSize = PaperKind.A4,
                Margins = new MarginSettings() { Top = 10 },
                Out = pdfFilePath,
            };

            // 传入配置
            if (configGlobalSettings != null)
            {
                configGlobalSettings(globalSettings);
            }

            
            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = globalSettings,
                Objects = {
                        new ObjectSettings()
                        {
                            Page = htmlFilePath,
                        },
                    }
            };            

            converter.Convert(doc);
        }
    }
}
