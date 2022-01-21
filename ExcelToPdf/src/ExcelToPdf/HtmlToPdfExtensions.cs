using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WkHtmlToPdfDotNet;
using WkHtmlToPdfDotNet.Contracts;

namespace CommonLibsCore.Pdf.Extensions
{
    /// <summary>
    /// Html to Pdf Extensions
    /// </summary>
    public static class HtmlToPdfExtensions
    {
        /// <summary>
        /// Add converter to DI
        /// </summary>
        /// <param name="services"></param>
        public static void AddHtmlToPdf(this IServiceCollection services)
        {            
            services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));
        }
    }
}
