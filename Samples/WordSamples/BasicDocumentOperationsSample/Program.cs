//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace BasicDocumentOperationsSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 基本文档操作示例");

            // 示例1: 应用程序基础属性操作
            Console.WriteLine("\n=== 示例1: 应用程序基础属性操作 ===");
            ApplicationPropertiesDemo();

            // 示例2: 文档集合管理
            Console.WriteLine("\n=== 示例2: 文档集合管理 ===");
            DocumentCollectionDemo();

            // 示例3: 活动文档和窗口管理
            Console.WriteLine("\n=== 示例3: 活动文档和窗口管理 ===");
            ActiveDocumentAndWindowDemo();

            // 示例4: 应用程序设置和选项
            Console.WriteLine("\n=== 示例4: 应用程序设置和选项 ===");
            ApplicationSettingsDemo();

            // 示例5: 文档保存和关闭操作
            Console.WriteLine("\n=== 示例5: 文档保存和关闭操作 ===");
            DocumentSaveAndCloseDemo();

            // 示例6: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例6: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 应用程序基础属性操作示例
        /// </summary>
        static void ApplicationPropertiesDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();

                // 设置应用程序窗口标题
                app.Caption = "我的文档编辑器 - 基础属性示例";
                Console.WriteLine($"设置窗口标题为: {app.Caption}");

                // 控制状态栏显示
                app.DisplayStatusBar = true;
                Console.WriteLine($"状态栏显示: {app.DisplayStatusBar}");

                // 控制滚动条显示
                app.DisplayScrollBars = true;
                Console.WriteLine($"滚动条显示: {app.DisplayScrollBars}");

                // 获取应用程序可用区域尺寸
                int width = app.UsableWidth;
                int height = app.UsableHeight;
                Console.WriteLine($"可用区域尺寸: {width} x {height} 磅");

                // 控制应用程序可见性
                app.Visible = true;
                Console.WriteLine("应用程序已设为可见");

                // 获取应用程序版本信息
                string version = app.Version;
                Console.WriteLine($"Word版本: {version}");

                Console.WriteLine("应用程序基础属性操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用程序基础属性操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文档集合管理示例
        /// </summary>
        static void DocumentCollectionDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();

                // 获取文档集合
                using var documents = app.Documents;
                Console.WriteLine($"初始文档数量: {documents.Count}");

                // 创建新文档
                using var newDoc = documents.Add();
                newDoc.Range().Text = "这是第一个文档\n创建时间: " + DateTime.Now.ToString();
                Console.WriteLine($"创建新文档后文档数量: {documents.Count}");

                // 再创建一个新文档
                using var newDoc2 = documents.Add();
                newDoc2.Range().Text = "这是第二个文档\n创建时间: " + DateTime.Now.ToString();
                Console.WriteLine($"再次创建新文档后文档数量: {documents.Count}");

                // 遍历所有文档
                Console.WriteLine("所有文档列表:");
                for (int i = 1; i <= documents.Count; i++)
                {
                    using var doc = documents[i];
                    Console.WriteLine($"  文档 {i}: {doc.Name}");
                }

                // 保存第一个文档
                string firstDocPath = Path.Combine(Path.GetTempPath(), "FirstDocument.docx");
                newDoc.SaveAs(firstDocPath);
                Console.WriteLine($"第一个文档已保存到: {firstDocPath}");

                Console.WriteLine("文档集合管理操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文档集合管理操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 活动文档和窗口管理示例
        /// </summary>
        static void ActiveDocumentAndWindowDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();

                // 创建多个文档
                using var doc1 = app.Documents.Add();
                doc1.Range().Text = "文档1内容";

                using var doc2 = app.Documents.Add();
                doc2.Range().Text = "文档2内容";

                // 获取活动文档
                using var activeDoc = app.ActiveDocument;
                if (activeDoc != null)
                {
                    Console.WriteLine($"当前活动文档: {activeDoc.Name}");
                }

                // 获取活动窗口
                using var activeWindow = app.ActiveWindow;
                if (activeWindow != null)
                {
                    Console.WriteLine($"当前活动窗口标题: {activeWindow.Caption}");
                }

                Console.WriteLine("活动文档和窗口管理操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"活动文档和窗口管理操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用程序设置和选项示例
        /// </summary>
        static void ApplicationSettingsDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();

                // 设置显示警告级别
                app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                Console.WriteLine("已禁用所有警告对话框");

                // 控制自动完成提示
                app.DisplayAutoCompleteTips = false;
                Console.WriteLine($"自动完成提示显示: {app.DisplayAutoCompleteTips}");

                // 控制屏幕提示显示
                app.DisplayScreenTips = true;
                Console.WriteLine($"屏幕提示显示: {app.DisplayScreenTips}");

                // 设置取消键处理方式
                app.EnableCancelKey = WdEnableCancelKey.wdCancelDisabled;
                Console.WriteLine("已启用取消键处理");

                // 控制语言检查
                app.CheckLanguage = true;
                Console.WriteLine($"语言检查启用: {app.CheckLanguage}");

                // 控制屏幕更新
                app.ScreenUpdating = false;
                Console.WriteLine("已关闭屏幕更新以提高性能");

                // 设置窗口状态
                app.WindowState = WdWindowState.wdWindowStateMaximize;
                Console.WriteLine("窗口已最大化");

                Console.WriteLine("应用程序设置和选项操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用程序设置和选项操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文档保存和关闭操作示例
        /// </summary>
        static void DocumentSaveAndCloseDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 添加文档内容
                document.Range().Text = "文档保存和关闭示例\n\n创建时间: " + DateTime.Now.ToString();

                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "BasicDocumentOperations");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 保存文档到不同格式
                string docxPath = Path.Combine(tempDirectory, "SaveExample.docx");
                document.SaveAs(docxPath);
                Console.WriteLine($"文档已保存为DOCX格式: {docxPath}");

                string docPath = Path.Combine(tempDirectory, "SaveExample.doc");
                document.SaveAs(docPath, WdSaveFormat.wdFormatDocument);
                Console.WriteLine($"文档已保存为DOC格式: {docPath}");

                string pdfPath = Path.Combine(tempDirectory, "SaveExample.pdf");
                document.SaveAs(pdfPath, WdSaveFormat.wdFormatPDF);
                Console.WriteLine($"文档已保存为PDF格式: {pdfPath}");

                // 保存副本
                string copyPath = Path.Combine(tempDirectory, "SaveCopyExample.docx");
                document.SaveAs(copyPath);
                Console.WriteLine($"文档副本已保存: {copyPath}");

                Console.WriteLine("文档保存和关闭操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文档保存和关闭操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                Console.WriteLine("使用WordDocumentManager辅助类进行完整操作:");

                // 创建文档管理器实例
                var documentManager = new WordDocumentManager();

                // 创建新文档
                string documentPath = documentManager.CreateNewDocument("完整示例文档\n\n这是使用辅助类创建的文档。");
                Console.WriteLine($"使用辅助类创建的文档: {documentPath}");

                // 从模板创建文档
                string templatePath = documentPath; // 使用刚才创建的文档作为模板
                string fromTemplatePath = documentManager.CreateDocumentFromTemplate(templatePath, "从模板创建的文档\n\n这是基于模板创建的文档。");
                Console.WriteLine($"从模板创建的文档: {fromTemplatePath}");

                // 打开现有文档
                string openedDocumentPath = documentManager.OpenAndModifyDocument(documentPath, "\n\n已修改的内容");
                Console.WriteLine($"修改后的文档: {openedDocumentPath}");

                // 批量创建文档
                var batchPaths = documentManager.CreateBatchDocuments(3);
                Console.WriteLine($"批量创建了 {batchPaths.Count} 个文档");

                Console.WriteLine("使用辅助类的完整示例操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例操作出错: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Word文档管理器辅助类
    /// </summary>
    public class WordDocumentManager
    {
        /// <summary>
        /// 创建新文档
        /// </summary>
        /// <param name="content">文档内容</param>
        /// <returns>文档路径</returns>
        public string CreateNewDocument(string content)
        {
            using var app = WordFactory.BlankDocument();
            using var document = app.ActiveDocument;

            // 添加内容
            document.Range().Text = content;

            // 保存文档
            string fileName = $"Document_{Guid.NewGuid()}.docx";
            string filePath = Path.Combine(Path.GetTempPath(), fileName);
            document.SaveAs(filePath);

            return filePath;
        }

        /// <summary>
        /// 从模板创建文档
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="additionalContent">附加内容</param>
        /// <returns>文档路径</returns>
        public string CreateDocumentFromTemplate(string templatePath, string additionalContent)
        {
            using var app = WordFactory.CreateFrom(templatePath);
            var document = app.ActiveDocument;

            // 添加附加内容
            document.Range().Text += additionalContent;

            // 保存文档
            string fileName = $"DocumentFromTemplate_{Guid.NewGuid()}.docx";
            string filePath = Path.Combine(Path.GetTempPath(), fileName);
            document.SaveAs(filePath);

            return filePath;
        }

        /// <summary>
        /// 打开并修改现有文档
        /// </summary>
        /// <param name="documentPath">文档路径</param>
        /// <param name="additionalContent">附加内容</param>
        /// <returns>文档路径</returns>
        public string OpenAndModifyDocument(string documentPath, string additionalContent)
        {
            using var app = WordFactory.Open(documentPath);
            using var document = app.ActiveDocument;

            // 添加附加内容
            document.Range().Text += additionalContent;

            // 保存文档
            document.Save();

            return documentPath;
        }

        /// <summary>
        /// 批量创建文档
        /// </summary>
        /// <param name="count">文档数量</param>
        /// <returns>文档路径列表</returns>
        public List<string> CreateBatchDocuments(int count)
        {
            var documentPaths = new List<string>();

            for (int i = 1; i <= count; i++)
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 添加内容
                document.Range().Text = $"批量创建的文档 {i}\n\n创建时间: {DateTime.Now}";

                // 保存文档
                string fileName = $"BatchDocument_{i}_{Guid.NewGuid()}.docx";
                string filePath = Path.Combine(Path.GetTempPath(), fileName);
                document.SaveAs(filePath);

                documentPaths.Add(filePath);
            }

            return documentPaths;
        }
    }
}