//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;
using System.Text;

namespace TaskPanesAndDialogsSample
{
    /// <summary>
    /// 用户交互处理器类
    /// </summary>
    public class UserInteractionHandler
    {
        private readonly IWordApplication _application;
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        /// <param name="document">Word文档对象</param>
        public UserInteractionHandler(IWordApplication? Application, IWordDocument document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 处理段落格式化交互
        /// </summary>
        /// <returns>是否处理成功</returns>
        public bool HandleParagraphFormattingInteraction()
        {
            try
            {
                Console.WriteLine("开始处理段落格式化交互...");

                // 显示段落对话框
                var dialogManager = new DialogManager(_application, _document);
                bool result = dialogManager.ShowParagraphDialog();

                if (result)
                {
                    Console.WriteLine("段落格式化操作已完成");
                    return true;
                }
                else
                {
                    Console.WriteLine("用户取消了段落格式化操作");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理段落格式化交互时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 处理页面设置交互
        /// </summary>
        /// <returns>是否处理成功</returns>
        public bool HandlePageSetupInteraction()
        {
            try
            {
                Console.WriteLine("开始处理页面设置交互...");

                // 显示页面设置对话框
                var dialogManager = new DialogManager(_application, _document);
                var pageSettings = dialogManager.ShowCustomPageSetupDialog();

                if (pageSettings.IsApplied)
                {
                    Console.WriteLine("页面设置已应用");
                    return true;
                }
                else
                {
                    Console.WriteLine("用户取消了页面设置操作");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理页面设置交互时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 处理查找替换交互
        /// </summary>
        /// <returns>是否处理成功</returns>
        public bool HandleFindReplaceInteraction()
        {
            try
            {
                Console.WriteLine("开始处理查找替换交互...");

                // 显示查找对话框
                var dialogManager = new DialogManager(_application, _document);
                bool result = dialogManager.ShowFindDialog();

                if (result)
                {
                    Console.WriteLine("查找替换操作已完成");
                    return true;
                }
                else
                {
                    Console.WriteLine("用户取消了查找替换操作");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理查找替换交互时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 处理文档属性交互
        /// </summary>
        /// <returns>文档属性信息</returns>
        public DocumentProperties HandleDocumentPropertiesInteraction()
        {
            var properties = new DocumentProperties();

            try
            {
                Console.WriteLine("开始处理文档属性交互...");

                // 获取当前文档属性
                properties.Title = _document.Title ?? "";
                properties.Author = _document.Author ?? "";
                properties.Subject = _document.Subject ?? "";
                properties.Keywords = _document.Keywords ?? "";

                Console.WriteLine("当前文档属性:");
                Console.WriteLine($"  标题: {properties.Title}");
                Console.WriteLine($"  作者: {properties.Author}");
                Console.WriteLine($"  主题: {properties.Subject}");
                Console.WriteLine($"  关键词: {properties.Keywords}");

                // 模拟用户修改属性（在实际应用中会显示对话框）
                Console.WriteLine("模拟用户修改文档属性...");
                properties.Title = "更新后的文档标题";
                properties.Author = "更新后的作者";
                properties.Subject = "更新后的主题";
                properties.Keywords = "更新后的关键词";

                // 应用更改
                _document.Title = properties.Title;
                _document.Author = properties.Author;
                _document.Subject = properties.Subject;
                _document.Keywords = properties.Keywords;

                properties.IsUpdated = true;
                Console.WriteLine("文档属性已更新");

                return properties;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理文档属性交互时出错: {ex.Message}");
                properties.ErrorMessage = ex.Message;
                return properties;
            }
        }

        /// <summary>
        /// 处理文件操作交互
        /// </summary>
        /// <param name="operationType">操作类型</param>
        /// <returns>文件操作结果</returns>
        public FileOperationResult HandleFileOperationInteraction(FileOperationType operationType)
        {
            var result = new FileOperationResult();

            try
            {
                Console.WriteLine($"开始处理文件操作交互: {operationType}");

                switch (operationType)
                {
                    case FileOperationType.Open:
                        // 模拟打开文件操作
                        Console.WriteLine("模拟打开文件对话框...");
                        result.FilePath = @"C:\temp\示例文档.docx";
                        result.IsSuccess = true;
                        result.Message = "文件打开操作已完成";
                        break;

                    case FileOperationType.Save:
                        // 模拟保存文件操作
                        Console.WriteLine("模拟保存文件对话框...");
                        result.FilePath = @"C:\temp\保存的文档.docx";
                        _document.Save(result.FilePath);
                        result.IsSuccess = true;
                        result.Message = "文件保存操作已完成";
                        break;

                    case FileOperationType.SaveAs:
                        // 模拟另存为操作
                        Console.WriteLine("模拟另存为对话框...");
                        result.FilePath = @"C:\temp\另存为文档.docx";
                        _document.Save(result.FilePath);
                        result.IsSuccess = true;
                        result.Message = "文件另存为操作已完成";
                        break;

                    default:
                        result.Message = "不支持的文件操作类型";
                        break;
                }

                Console.WriteLine(result.Message);
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理文件操作交互时出错: {ex.Message}");
                result.IsSuccess = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
        }

        /// <summary>
        /// 生成用户交互报告
        /// </summary>
        /// <param name="interactions">交互记录列表</param>
        /// <returns>交互报告</returns>
        public string GenerateInteractionReport(List<UserInteractionRecord> interactions)
        {
            var reportBuilder = new StringBuilder();
            reportBuilder.AppendLine("=== 用户交互报告 ===");
            reportBuilder.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            reportBuilder.AppendLine();

            if (interactions == null || interactions.Count == 0)
            {
                reportBuilder.AppendLine("没有用户交互记录");
                return reportBuilder.ToString();
            }

            reportBuilder.AppendLine($"总交互次数: {interactions.Count}");
            reportBuilder.AppendLine();

            foreach (var interaction in interactions)
            {
                reportBuilder.AppendLine($"时间: {interaction.Timestamp:yyyy-MM-dd HH:mm:ss}");
                reportBuilder.AppendLine($"类型: {interaction.InteractionType}");
                reportBuilder.AppendLine($"描述: {interaction.Description}");
                reportBuilder.AppendLine($"结果: {interaction.Result}");
                reportBuilder.AppendLine(new string('-', 30));
            }

            reportBuilder.AppendLine("==================");

            return reportBuilder.ToString();
        }

        /// <summary>
        /// 创建交互式文档工具
        /// </summary>
        /// <returns>是否创建成功</returns>
        public bool CreateInteractiveDocumentTool()
        {
            try
            {
                Console.WriteLine("开始创建交互式文档工具...");

                // 创建示例文档内容
                var range = _document.Range();
                range.Text = "交互式文档工具演示\n\n" +
                            "此文档演示了如何处理用户交互操作。\n\n" +
                            "支持的交互操作包括：\n" +
                            "1. 字体格式化\n" +
                            "2. 段落格式化\n" +
                            "3. 页面设置\n" +
                            "4. 查找替换\n" +
                            "5. 文档属性管理\n" +
                            "6. 文件操作\n\n" +
                            "通过这些交互功能，用户可以方便地操作和格式化文档。";

                // 格式化标题
                var titleRange = _document.Range(0, 12);
                titleRange.Font.Size = 16;
                titleRange.Font.Bold = true;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 格式化列表
                var listStart = range.Text.IndexOf("支持的交互操作包括：");
                var listEnd = range.Text.IndexOf("通过这些交互功能");
                if (listStart > 0 && listEnd > listStart)
                {
                    var listRange = _document.Range(listStart, listEnd);
                    listRange.ListFormat.ApplyBulletDefault();
                }

                Console.WriteLine("交互式文档工具已创建");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建交互式文档工具时出错: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// 文档属性类
    /// </summary>
    public class DocumentProperties
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// 主题
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// 关键词
        /// </summary>
        public string Keywords { get; set; }

        /// <summary>
        /// 是否已更新
        /// </summary>
        public bool IsUpdated { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 文件操作结果类
    /// </summary>
    public class FileOperationResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 文件路径
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 文件操作类型枚举
    /// </summary>
    public enum FileOperationType
    {
        /// <summary>
        /// 打开文件
        /// </summary>
        Open,

        /// <summary>
        /// 保存文件
        /// </summary>
        Save,

        /// <summary>
        /// 另存为
        /// </summary>
        SaveAs
    }

    /// <summary>
    /// 用户交互记录类
    /// </summary>
    public class UserInteractionRecord
    {
        /// <summary>
        /// 时间戳
        /// </summary>
        public DateTime Timestamp { get; set; }

        /// <summary>
        /// 交互类型
        /// </summary>
        public string InteractionType { get; set; }

        /// <summary>
        /// 描述
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 结果
        /// </summary>
        public string Result { get; set; }
    }
}