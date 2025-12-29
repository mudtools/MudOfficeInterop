//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace WordFactorySample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - WordFactory 示例");

            // 示例1: 创建空白文档
            Console.WriteLine("\n=== 示例1: 创建空白文档 ===");
            CreateBlankDocument();

            // 示例2: 基于模板创建文档
            Console.WriteLine("\n=== 示例2: 基于模板创建文档 ===");
            CreateFromTemplate();

            // 示例3: 打开现有文档
            Console.WriteLine("\n=== 示例3: 打开现有文档 ===");
            OpenExistingDocument();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 创建空白文档示例
        /// </summary>
        static void CreateBlankDocument()
        {
            try
            {
                // 创建 Word 应用程序实例和空白文档
                using var app = WordFactory.BlankDocument();
                app.Visible = true;

                // 获取活动文档
                var document = app.ActiveDocument;

                // 添加内容
                var range = document.Range();
                range.Text = "这是使用 MudTools.OfficeInterop.Word 创建的空白文档示例。\n\n";
                range.InsertAfter("文档创建时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                Console.WriteLine("空白文档创建成功");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建空白文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 基于模板创建文档示例
        /// </summary>
        static void CreateFromTemplate()
        {
            try
            {
                // 创建临时模板文件
                string templatePath = Path.Combine(Path.GetTempPath(), "SampleTemplate.dotx");
                CreateSampleTemplate(templatePath);

                // 基于模板创建文档
                using var app = WordFactory.CreateFrom(templatePath);
                app.Visible = true;

                var document = app.ActiveDocument;

                // 替换模板中的占位符
                var selection = app.Selection;
                selection.Find.Text = "{REPORT_TITLE}";
                selection.Find.Replacement.Text = "季度销售报告";
                selection.Find.Execute(replace: WdReplace.wdReplaceAll);

                selection.Find.Text = "{DATE}";
                selection.Find.Replacement.Text = DateTime.Now.ToString("yyyy年MM月dd日");
                selection.Find.Execute(replace: WdReplace.wdReplaceAll);

                Console.WriteLine("基于模板创建文档成功");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"基于模板创建文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 打开现有文档示例
        /// </summary>
        static void OpenExistingDocument()
        {
            try
            {
                // 创建一个临时文档作为示例
                string tempDocPath = Path.Combine(Path.GetTempPath(), "SampleDocument.docx");
                CreateSampleDocument(tempDocPath);

                // 打开现有文档
                using var app = WordFactory.Open(tempDocPath);
                app.Visible = true;

                var document = app.ActiveDocument;

                // 在文档末尾追加内容
                document.Range().InsertAfter($"\n\n文档打开时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

                Console.WriteLine("打开现有文档成功");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打开现有文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建示例模板文件
        /// </summary>
        /// <param name="templatePath">模板文件路径</param>
        static void CreateSampleTemplate(string templatePath)
        {
            using var app = WordFactory.BlankDocument();
            var document = app.ActiveDocument;

            // 添加模板内容
            var range = document.Range();
            range.Text = "{REPORT_TITLE}\n\n";
            range.InsertAfter("报告日期: {DATE}\n\n");
            range.InsertAfter("这是一个示例模板文件。\n\n");
            range.InsertAfter("模板创建时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            // 保存为模板
            document.SaveAs(templatePath, WdSaveFormat.wdFormatTemplate);
        }

        /// <summary>
        /// 创建示例文档文件
        /// </summary>
        /// <param name="docPath">文档文件路径</param>
        static void CreateSampleDocument(string docPath)
        {
            using var app = WordFactory.BlankDocument();
            var document = app.ActiveDocument;

            // 添加文档内容
            var range = document.Range();
            range.Text = "示例文档\n\n";
            range.InsertAfter("这是一个示例文档文件。\n\n");
            range.InsertAfter("文档创建时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            // 保存文档
            document.SaveAs(docPath);
        }
    }
}