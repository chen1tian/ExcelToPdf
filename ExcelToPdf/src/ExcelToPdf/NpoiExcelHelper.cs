using NPOI.HSSF.UserModel;
using NPOI.SS.Converter;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

namespace ExcelToPdf
{
    /// <summary>
    /// 基于Npoi的Excel处理类
    /// </summary>
    public class NpoiExcelHelper
    {
        /// <summary>
        /// excel转html
        /// <paramref name="excelPath">excel文件路径</paramref>
        /// <paramref name="htmlFilePath">html文件地址</paramref>
        /// <paramref name="removeSheetName">是否移除生成后的Sheet名称</paramref>
        /// <paramref name="options">配置转换器</paramref>
        /// <paramref name="afterProcess">处理完成后的动作</paramref>
        /// </summary>
        public static void ExcelToHtml(string excelPath, string htmlFilePath, bool removeSheetName = true, Action<ExcelToHtmlConverter> options = null, Action<ExcelToHtmlConverter> afterProcess = null)
        {
            IWorkbook workbook;
            FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.ReadWrite);

            string fileExt = Path.GetExtension(excelPath).ToLower();
            if (fileExt == ".xlsx")
            {
                workbook = new XSSFWorkbook(fs);
            }
            else if (fileExt == ".xls")
            {
                workbook = new HSSFWorkbook(fs);
            }
            else
            {
                throw new Exception($"不支持文件格式{fileExt}");
            }

            ExcelToHtml(workbook, htmlFilePath, removeSheetName, options, afterProcess);
            workbook.Close();
            fs.Close();
        }

        /// <summary>
        /// excel转html
        /// <paramref name="workbook">工作簿对象</paramref>
        /// <paramref name="htmlFilePath">html文件地址</paramref>
        /// <paramref name="removeSheetName">是否移除生成后的Sheet名称</paramref>
        /// <paramref name="options">配置转换器</paramref>
        /// <paramref name="afterProcess">处理完成后的动作</paramref>
        /// </summary>
        public static void ExcelToHtml(IWorkbook workbook, string htmlFilePath, bool removeSheetName = true, Action<ExcelToHtmlConverter> options = null, Action<ExcelToHtmlConverter> afterProcess = null)
        {
            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter();

            // 配置转换器
            if (options != null)
            {
                options(excelToHtmlConverter);
            }
            else
            {
                // 设置输出参数			
                excelToHtmlConverter.OutputColumnHeaders = false;
                excelToHtmlConverter.OutputRowNumbers = false;
            }

            // 处理
            excelToHtmlConverter.ProcessWorkbook(workbook);

            // 如果要移除SheetName
            if (removeSheetName)
            {
                // 去掉html中的sheet名称
                var clearHtml = Regex.Replace(excelToHtmlConverter.Document.InnerXml, "<h2>.*</h2>", "");
                excelToHtmlConverter.Document.InnerXml = clearHtml;
            }

            // 处理后工作
            if (afterProcess != null)
            {
                afterProcess(excelToHtmlConverter);
            }

            //添加表格样式 修复linux环境下转pdf后换行的问题
            excelToHtmlConverter.Document.InnerXml =
                excelToHtmlConverter.Document.InnerXml.Insert(
                    excelToHtmlConverter.Document.InnerXml.IndexOf("</style>", 0) - 1,
                    @" table, tr, td, th, tbody, thead, tfoot { page-break-inside: avoid!important; }"
                );

            // 输出
            excelToHtmlConverter.Document.Save(htmlFilePath);
        }
    }
}