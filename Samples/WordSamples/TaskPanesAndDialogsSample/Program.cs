//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace TaskPanesAndDialogsSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 任务窗格和对话框示例");

            // 示例1: 任务窗格操作
            Console.WriteLine("\n=== 示例1: 任务窗格操作 ===");
            TaskPaneOperationsDemo();

            // 示例2: 对话框操作
            Console.WriteLine("\n=== 示例2: 对话框操作 ===");
            DialogOperationsDemo();

            // 示例3: 用户交互处理
            Console.WriteLine("\n=== 示例3: 用户交互处理 ===");
            UserInteractionDemo();

            // 示例4: 实际应用示例
            Console.WriteLine("\n=== 示例4: 实际应用示例 ===");
            RealWorldApplicationDemo();

            // 示例5: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例5: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 任务窗格操作示例
        /// </summary>
        static void TaskPaneOperationsDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                var taskPaneManager = new TaskPaneManager(app);

                // 获取任务窗格信息
                Console.WriteLine("1. 获取任务窗格信息");
                var taskPaneInfo = taskPaneManager.GetTaskPaneInfo();
                Console.WriteLine($"任务窗格数量: {taskPaneInfo.TaskPaneCount}");
                Console.WriteLine($"是否支持任务窗格: {taskPaneInfo.IsTaskPaneSupported}");

                // 创建任务窗格XML定义
                Console.WriteLine("\n2. 创建任务窗格XML定义");
                string xmlDefinition = taskPaneManager.CreateTaskPaneXml(
                    "tabCustom",
                    "自定义工具",
                    "grpTaskPane",
                    "任务窗格",
                    "btnShowTaskPane",
                    "显示任务窗格");
                Console.WriteLine("XML定义已生成");

                // 生成任务窗格用户控件代码
                Console.WriteLine("\n3. 生成任务窗格用户控件代码");
                string userControlCode = taskPaneManager.GenerateTaskPaneUserControlCode();
                Console.WriteLine("用户控件代码示例已生成");

                // 生成VSTO插件任务窗格实现代码
                Console.WriteLine("\n4. 生成VSTO插件任务窗格实现代码");
                string vstoCode = taskPaneManager.GenerateVstoTaskPaneImplementation();
                Console.WriteLine("VSTO插件代码示例已生成");

                // 生成任务窗格最佳实践指南
                Console.WriteLine("\n5. 生成任务窗格最佳实践指南");
                string bestPractices = taskPaneManager.GenerateTaskPaneBestPractices();
                Console.WriteLine("最佳实践指南已生成");

                Console.WriteLine("\n任务窗格操作示例演示完成");
                Console.WriteLine("注意：完整的任务窗格功能需要在VSTO插件环境中实现");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"任务窗格操作示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 对话框操作示例
        /// </summary>
        static void DialogOperationsDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                var document = app.ActiveDocument;
                var dialogManager = new DialogManager(app, document);

                // 获取所有对话框信息
                Console.WriteLine("1. 获取所有对话框信息");
                var dialogInfos = dialogManager.GetAllDialogsInfo();
                Console.WriteLine($"找到 {dialogInfos.Count} 个对话框");

                // 显示字体对话框
                Console.WriteLine("\n2. 显示字体对话框");
                // 注意：在实际运行中，这会显示字体对话框，但在这里我们只演示代码
                Console.WriteLine("字体对话框演示代码已准备");

                // 显示段落对话框
                Console.WriteLine("\n3. 显示段落对话框");
                Console.WriteLine("段落对话框演示代码已准备");

                // 显示页面设置对话框
                Console.WriteLine("\n4. 显示页面设置对话框");
                Console.WriteLine("页面设置对话框演示代码已准备");

                // 显示查找对话框
                Console.WriteLine("\n5. 显示查找对话框");
                Console.WriteLine("查找对话框演示代码已准备");

                // 生成对话框交互最佳实践指南
                Console.WriteLine("\n6. 生成对话框交互最佳实践指南");
                string bestPractices = dialogManager.GenerateDialogBestPractices();
                Console.WriteLine("最佳实践指南已生成");

                Console.WriteLine("\n对话框操作示例演示完成");
                Console.WriteLine("注意：在实际运行中，对话框会显示在屏幕上");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"对话框操作示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 用户交互处理示例
        /// </summary>
        static void UserInteractionDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                var document = app.ActiveDocument;
                var interactionHandler = new UserInteractionHandler(app, document);

                // 创建交互式文档工具
                Console.WriteLine("1. 创建交互式文档工具");
                bool toolCreated = interactionHandler.CreateInteractiveDocumentTool();
                Console.WriteLine($"工具创建结果: {toolCreated}");

                // 处理字体格式化交互
                Console.WriteLine("\n2. 处理字体格式化交互");
                Console.WriteLine("字体格式化交互演示代码已准备");

                // 处理段落格式化交互
                Console.WriteLine("\n3. 处理段落格式化交互");
                Console.WriteLine("段落格式化交互演示代码已准备");

                // 处理页面设置交互
                Console.WriteLine("\n4. 处理页面设置交互");
                Console.WriteLine("页面设置交互演示代码已准备");

                // 处理查找替换交互
                Console.WriteLine("\n5. 处理查找替换交互");
                Console.WriteLine("查找替换交互演示代码已准备");

                // 处理文档属性交互
                Console.WriteLine("\n6. 处理文档属性交互");
                var properties = interactionHandler.HandleDocumentPropertiesInteraction();
                Console.WriteLine($"文档属性更新结果: {properties.IsUpdated}");

                // 处理文件操作交互
                Console.WriteLine("\n7. 处理文件操作交互");
                var openResult = interactionHandler.HandleFileOperationInteraction(FileOperationType.Open);
                Console.WriteLine($"打开文件操作结果: {openResult.IsSuccess}");

                var saveResult = interactionHandler.HandleFileOperationInteraction(FileOperationType.Save);
                Console.WriteLine($"保存文件操作结果: {saveResult.IsSuccess}");

                Console.WriteLine("\n用户交互处理示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"用户交互处理示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例
        /// </summary>
        static void RealWorldApplicationDemo()
        {
            try
            {
                Console.WriteLine("=== 任务窗格和对话框系统演示 ===");
                Console.WriteLine();

                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "TaskPanesAndDialogs");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                using var app = WordFactory.BlankDocument();
                var document = app.ActiveDocument;

                // 步骤1: 创建支持任务窗格和对话框的文档
                Console.WriteLine("步骤1: 创建支持任务窗格和对话框的文档");
                document.Range().Text = "任务窗格和对话框支持文档\n\n" +
                                      "此文档演示了如何为Word开发支持自定义任务窗格和对话框交互的插件。\n\n" +
                                      "主要特性包括：\n" +
                                      "1. 自定义任务窗格\n" +
                                      "2. 对话框操作\n" +
                                      "3. 用户交互处理\n" +
                                      "4. 实时格式化\n\n" +
                                      "完整实现需要在VSTO插件环境中进行。";

                // 格式化标题
                var titleRange = document.Range(0, 15);
                titleRange.Font.Size = 16;
                titleRange.Font.Bold = true;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 格式化列表
                var listStart = document.Range().Text.IndexOf("主要特性包括：");
                var listEnd = document.Range().Text.IndexOf("完整实现需要在VSTO插件环境中进行。");
                if (listStart > 0 && listEnd > listStart)
                {
                    var listRange = document.Range(listStart, listEnd);
                    listRange.ListFormat.ApplyBulletDefault();
                }

                // 保存文档
                string filePath = Path.Combine(tempDirectory, "DialogSupportingDocument.docx");
                document.Save(filePath);

                Console.WriteLine($"支持对话框的文档已创建: {filePath}");
                Console.WriteLine();

                // 步骤2: 演示任务窗格概念
                Console.WriteLine("步骤2: 演示任务窗格概念");
                var taskPaneManager = new TaskPaneManager(app);
                string xmlDefinition = taskPaneManager.CreateTaskPaneXml(
                    "tabCustom",
                    "自定义工具",
                    "grpTaskPane",
                    "任务窗格",
                    "btnShowTaskPane",
                    "显示任务窗格");
                Console.WriteLine("任务窗格XML定义已生成");
                Console.WriteLine();

                // 步骤3: 演示对话框交互
                Console.WriteLine("步骤3: 演示对话框交互");
                var dialogManager = new DialogManager(app, document);
                var dialogInfos = dialogManager.GetAllDialogsInfo();
                Console.WriteLine($"可用对话框数量: {dialogInfos.Count}");
                Console.WriteLine();

                // 步骤4: 处理用户交互
                Console.WriteLine("步骤4: 处理用户交互");
                var interactionHandler = new UserInteractionHandler(app, document);
                var properties = interactionHandler.HandleDocumentPropertiesInteraction();
                Console.WriteLine($"文档属性处理结果: {properties.IsUpdated}");

                var fileOperationResult = interactionHandler.HandleFileOperationInteraction(FileOperationType.SaveAs);
                Console.WriteLine($"文件操作处理结果: {fileOperationResult.IsSuccess}");
                Console.WriteLine();

                // 步骤5: 生成报告
                Console.WriteLine("步骤5: 生成报告");
                var interactions = new List<UserInteractionRecord>
                {
                    new UserInteractionRecord
                    {
                        Timestamp = DateTime.Now,
                        InteractionType = "文档创建",
                        Description = "创建支持任务窗格和对话框的文档",
                        Result = "成功"
                    },
                    new UserInteractionRecord
                    {
                        Timestamp = DateTime.Now.AddSeconds(1),
                        InteractionType = "任务窗格",
                        Description = "生成任务窗格XML定义",
                        Result = "成功"
                    },
                    new UserInteractionRecord
                    {
                        Timestamp = DateTime.Now.AddSeconds(2),
                        InteractionType = "对话框",
                        Description = "获取可用对话框信息",
                        Result = "成功"
                    }
                };

                string report = interactionHandler.GenerateInteractionReport(interactions);
                Console.WriteLine("用户交互报告已生成");
                Console.WriteLine();

                Console.WriteLine("任务窗格和对话框系统演示完成！");
                Console.WriteLine($"生成的文档位于 {filePath}");
                Console.WriteLine("注意：完整的任务窗格和对话框功能需要在VSTO插件环境中实现");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"实际应用示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "TaskPanesAndDialogs");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                using var app = WordFactory.BlankDocument();
                var document = app.ActiveDocument;

                // 1. 任务窗格管理
                Console.WriteLine("1. 任务窗格管理");
                var taskPaneManager = new TaskPaneManager(app);
                var taskPaneInfo = taskPaneManager.GetTaskPaneInfo();
                Console.WriteLine($"任务窗格信息: 数量={taskPaneInfo.TaskPaneCount}, 支持={taskPaneInfo.IsTaskPaneSupported}");

                string xmlDefinition = taskPaneManager.CreateTaskPaneXml(
                    "tabComplete", "完整示例", "grpComplete", "完整功能", "btnComplete", "完整功能");
                Console.WriteLine("任务窗格XML定义已生成");

                string userControlCode = taskPaneManager.GenerateTaskPaneUserControlCode();
                Console.WriteLine("任务窗格用户控件代码已生成");
                Console.WriteLine();

                // 2. 对话框管理
                Console.WriteLine("2. 对话框管理");
                var dialogManager = new DialogManager(app, document);
                var dialogInfos = dialogManager.GetAllDialogsInfo();
                Console.WriteLine($"找到 {dialogInfos.Count} 个对话框");

                // 演示各种对话框操作
                Console.WriteLine("对话框操作演示代码已准备");
                Console.WriteLine();

                // 3. 用户交互处理
                Console.WriteLine("3. 用户交互处理");
                var interactionHandler = new UserInteractionHandler(app, document);
                bool toolCreated = interactionHandler.CreateInteractiveDocumentTool();
                Console.WriteLine($"交互式工具创建结果: {toolCreated}");

                var properties = interactionHandler.HandleDocumentPropertiesInteraction();
                Console.WriteLine($"文档属性处理结果: {properties.IsUpdated}");

                // 4. 文件操作处理
                Console.WriteLine("4. 文件操作处理");
                var openResult = interactionHandler.HandleFileOperationInteraction(FileOperationType.Open);
                Console.WriteLine($"打开文件操作结果: {openResult.IsSuccess}");

                var saveResult = interactionHandler.HandleFileOperationInteraction(FileOperationType.Save);
                Console.WriteLine($"保存文件操作结果: {saveResult.IsSuccess}");

                // 5. 生成综合报告
                Console.WriteLine("5. 生成综合报告");
                var interactions = new List<UserInteractionRecord>
                {
                    new UserInteractionRecord
                    {
                        Timestamp = DateTime.Now,
                        InteractionType = "任务窗格",
                        Description = "创建任务窗格XML定义",
                        Result = "成功"
                    },
                    new UserInteractionRecord
                    {
                        Timestamp = DateTime.Now.AddSeconds(1),
                        InteractionType = "对话框",
                        Description = "获取对话框信息",
                        Result = "成功"
                    },
                    new UserInteractionRecord
                    {
                        Timestamp = DateTime.Now.AddSeconds(2),
                        InteractionType = "用户交互",
                        Description = "处理文档属性",
                        Result = "成功"
                    }
                };

                string interactionReport = interactionHandler.GenerateInteractionReport(interactions);
                Console.WriteLine("综合交互报告已生成");
                Console.WriteLine();

                // 6. 最佳实践指南
                Console.WriteLine("6. 最佳实践指南");
                string taskPaneBestPractices = taskPaneManager.GenerateTaskPaneBestPractices();
                Console.WriteLine("任务窗格最佳实践指南已生成");

                string dialogBestPractices = dialogManager.GenerateDialogBestPractices();
                Console.WriteLine("对话框最佳实践指南已生成");
                Console.WriteLine();

                Console.WriteLine("使用辅助类的完整示例演示完成");
                Console.WriteLine("注意：完整的实现需要在VSTO插件环境中进行");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例演示出错: {ex.Message}");
            }
        }
    }
}