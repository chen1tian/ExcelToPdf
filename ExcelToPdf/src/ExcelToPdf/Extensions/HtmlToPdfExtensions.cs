using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WkHtmlToPdfDotNet;
using WkHtmlToPdfDotNet.Contracts;

namespace ExcelToPdf.Extensions
{
    /// <summary>
    /// Html to Pdf Extensions
    /// </summary>
    public static class HtmlToPdfExtensions
    {
        /// <summary>
        /// add ExcelToPdfService to DI
        /// </summary>
        /// <param name="services"></param>
        /// <param name="tempDir"></param>
        public static void AddExcelToPdf(this IServiceCollection services, string tempDir)
        {
            var pdfConverter = new SynchronizedConverter(new PdfTools());
            services.AddSingleton(typeof(IConverter), pdfConverter);

            services.AddTransient<ExcelToPdfService>(services =>
            {
                return new ExcelToPdfService(pdfConverter, tempDir);
            });
        }
    }
}
