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
        /// <param name="docFactory"></param>
        public static void HtmlToPdf(this IConverter converter, string htmlFilePath, string pdfFilePath, Func<HtmlToPdfDocument> docFactory = null)
        {
            HtmlToPdfDocument doc = null;

            // 如果为配置，那么使用默认设置
            if (docFactory != null)
            {
                doc = docFactory();
            }
            else
            {
                doc = new HtmlToPdfDocument()
                {
                    GlobalSettings = {
                        ColorMode = ColorMode.Color,
                        Orientation = Orientation.Portrait,
                        PaperSize = PaperKind.A4,
                        Margins = new MarginSettings() { Top = 10 },
                        Out = pdfFilePath,
                    },
                    Objects = {
                        new ObjectSettings()
                        {
                            Page = htmlFilePath,
                        },
                    }
                };
            }

            converter.Convert(doc);
        }
    }
}
