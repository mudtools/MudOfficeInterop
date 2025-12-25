//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;
using System.Drawing;
using System.Text;

namespace DocumentAutomationProcessingSample
{
    /// <summary>
    /// 文档自动化工作流类
    /// </summary>
    public class DocumentAutomationWorkflow
    {
        /// <summary>
        /// 工作流配置类
        /// </summary>
        public class WorkflowConfiguration
        {
            /// <summary>
            /// 是否标准化格式
            /// </summary>
            public bool StandardizeFormat { get; set; } = true;

            /// <summary>
            /// 是否更新字段
            /// </summary>
            public bool UpdateFields { get; set; } = true;

            /// <summary>
            /// 是否添加页眉页脚
            /// </summary>
            public bool AddHeaderFooter { get; set; } = true;

            /// <summary>
            /// 是否转换为PDF
            /// </summary>
            public bool ConvertToPdf { get; set; } = false;

            /// <summary>
            /// 是否生成目录
            /// </summary>
            public bool GenerateTableOfContents { get; set; } = false;

            /// <summary>
            /// 水印文本
            /// </summary>
            public string WatermarkText { get; set; } = null;

            /// <summary>
            /// 是否添加水印
            /// </summary>
            public bool AddWatermark => !string.IsNullOrEmpty(WatermarkText);
        }

        /// <summary>
        /// 执行工作流
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="config">工作流配置</param>
        /// <returns>工作流执行结果</returns>
        public static WorkflowResult ExecuteWorkflow(
            string inputDirectory,
            string outputDirectory,
            WorkflowConfiguration config)
        {
            var result = new WorkflowResult();

            try
            {
                Console.WriteLine("=== 开始执行文档自动化工作流 ===");

                // 确保输出目录存在
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // 获取所有Word文档
                var docFiles = Directory.GetFiles(inputDirectory, "*.doc");
                var docxFiles = Directory.GetFiles(inputDirectory, "*.docx");
                var allFiles = docFiles.Concat(docxFiles).ToArray();

                Console.WriteLine($"找到 {allFiles.Length} 个文档需要处理");

                result.TotalFiles = allFiles.Length;
                result.ProcessedFiles = new List<string>();
                result.FailedFiles = new List<string>();

                foreach (var file in allFiles)
                {
                    try
                    {
                        Console.WriteLine($"\n正在处理: {Path.GetFileName(file)}");

                        // 执行工作流
                        ExecuteSingleDocumentWorkflow(file, outputDirectory, config);

                        result.ProcessedFiles.Add(file);
                        Console.WriteLine($"处理完成: {Path.GetFileName(file)}");
                    }
                    catch (Exception ex)
                    {
                        result.FailedFiles.Add(file);
                        Console.WriteLine($"处理 {Path.GetFileName(file)} 时出错: {ex.Message}");
                    }
                }

                Console.WriteLine($"\n=== 工作流执行完成 ===");
                Console.WriteLine($"成功处理: {result.ProcessedFiles.Count} 个文档");
                Console.WriteLine($"处理失败: {result.FailedFiles.Count} 个文档");
                Console.WriteLine($"总计处理: {allFiles.Length} 个文档");

                result.Success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"工作流执行过程中出错: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 执行单个文档工作流
        /// </summary>
        /// <param name="inputFilePath">输入文件路径</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="config">工作流配置</param>
        private static void ExecuteSingleDocumentWorkflow(
            string inputFilePath,
            string outputDirectory,
            WorkflowConfiguration config)
        {
            using var app = WordFactory.Open(inputFilePath);
            using var document = app.ActiveDocument;

            // 执行配置的处理步骤
            if (config.StandardizeFormat)
            {
                StandardizeDocumentFormat(document);
            }

            if (config.UpdateFields)
            {
                UpdateDocumentFields(document);
            }

            if (config.AddHeaderFooter)
            {
                AddHeaderFooter(document);
            }

            if (config.GenerateTableOfContents)
            {
                GenerateTableOfContents(document);
            }

            if (config.AddWatermark)
            {
                AddWatermark(document, config.WatermarkText);
            }

            // 生成输出文件路径
            var fileName = Path.GetFileNameWithoutExtension(inputFilePath);
            var outputFilePath = Path.Combine(outputDirectory, $"{fileName}_processed.docx");

            // 保存处理后的文档
            document.SaveAs(outputFilePath);

            // 如果需要转换为PDF
            if (config.ConvertToPdf)
            {
                var pdfOutputPath = Path.Combine(outputDirectory, $"{fileName}_processed.pdf");
                document.SaveAs(pdfOutputPath, WdSaveFormat.wdFormatPDF);
                Console.WriteLine("  - 已转换为PDF格式");
            }
        }

        /// <summary>
        /// 标准化文档格式
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private static void StandardizeDocumentFormat(IWordDocument document)
        {
            try
            {
                // 标准化字体
                using var range = document.Range();
                range.Font.Name = "宋体";
                range.Font.Size = 12;

                // 标准化段落格式
                foreach (var paragraph in document.Paragraphs)
                {
                    using (paragraph)
                    {
                        paragraph.Format.LineSpacing = 1.5f; // 1.5倍行距
                        paragraph.Format.SpaceAfter = 12;    // 段后间距
                    }
                }
                Console.WriteLine("  - 文档格式已标准化");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"标准化文档格式时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新文档字段
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private static void UpdateDocumentFields(IWordDocument document)
        {
            try
            {
                // 更新所有字段
                document.Range().Fields.Update();
                Console.WriteLine("  - 文档字段已更新");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"更新文档字段时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加页眉页脚
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private static void AddHeaderFooter(IWordDocument document)
        {
            try
            {
                // 添加页眉
                using var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = "公司文档";
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加页脚（包含页码）
                using var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
                footerRange.Text = " 第 页";
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                Console.WriteLine("  - 页眉页脚已添加");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加页眉页脚时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 生成目录
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private static void GenerateTableOfContents(IWordDocument document)
        {
            try
            {
                // 查找目录插入位置（通常在文档开头）s
                using var range = document.Range(0, 0);
                range.Text = "目录\n";
                range.Font.Size = 16;
                range.Font.Bold = true;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 插入目录
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                document.TablesOfContents.Add(range);

                Console.WriteLine("  - 目录已生成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成目录时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加水印
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="watermarkText">水印文本</param>
        private static void AddWatermark(IWordDocument document, string watermarkText)
        {
            try
            {
                // 在每个节中添加水印
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    using var section = document.Sections[i];
                    using var header = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    using var shapes = header.Range.ShapeRange.Parent as IWordShapes;
                    using var shape = shapes.AddTextEffect(
                        MsoPresetTextEffect.msoTextEffect1,
                        watermarkText,
                        "Arial",
                        100,
                        false,
                        false,
                        0,
                        0);

                    // 设置水印属性
                    shape.Fill.Visible = true;
                    shape.Fill.Solid();
                    shape.Fill.ForeColor.RGB = Color.Gray;
                    shape.Line.Visible = false;
                    shape.Rotation = 315; // 斜角
                    shape.WrapFormat.AllowOverlap = true;
                    shape.WrapFormat.Type = WdWrapType.wdWrapNone;
                    shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                    shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
                    shape.LeftPosition = WdShapePosition.wdShapeCenter;
                    shape.TopPosition = WdShapePosition.wdShapeCenter;
                }

                Console.WriteLine("  - 水印已添加");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加水印时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行高级工作流
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="config">工作流配置</param>
        /// <returns>工作流执行结果</returns>
        public static async Task<WorkflowResult> ExecuteAdvancedWorkflowAsync(
            string inputDirectory,
            string outputDirectory,
            WorkflowConfiguration config)
        {
            return await Task.Run(() => ExecuteWorkflow(inputDirectory, outputDirectory, config));
        }

        /// <summary>
        /// 创建自定义工作流
        /// </summary>
        /// <param name="name">工作流名称</param>
        /// <param name="steps">工作流步骤</param>
        /// <returns>自定义工作流</returns>
        public static CustomWorkflow CreateCustomWorkflow(string name, List<WorkflowStep> steps)
        {
            return new CustomWorkflow
            {
                Name = name,
                Steps = steps
            };
        }

        /// <summary>
        /// 执行自定义工作流
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="customWorkflow">自定义工作流</param>
        /// <returns>工作流执行结果</returns>
        public static WorkflowResult ExecuteCustomWorkflow(
            string inputDirectory,
            string outputDirectory,
            CustomWorkflow customWorkflow)
        {
            var result = new WorkflowResult();

            try
            {
                Console.WriteLine($"=== 开始执行自定义工作流: {customWorkflow.Name} ===");

                // 确保输出目录存在
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // 获取所有Word文档
                var docFiles = Directory.GetFiles(inputDirectory, "*.doc");
                var docxFiles = Directory.GetFiles(inputDirectory, "*.docx");
                var allFiles = docFiles.Concat(docxFiles).ToArray();

                Console.WriteLine($"找到 {allFiles.Length} 个文档需要处理");

                result.TotalFiles = allFiles.Length;
                result.ProcessedFiles = new List<string>();
                result.FailedFiles = new List<string>();

                foreach (var file in allFiles)
                {
                    try
                    {
                        Console.WriteLine($"\n正在处理: {Path.GetFileName(file)}");

                        // 执行自定义工作流
                        ExecuteCustomDocumentWorkflow(file, outputDirectory, customWorkflow);

                        result.ProcessedFiles.Add(file);
                        Console.WriteLine($"处理完成: {Path.GetFileName(file)}");
                    }
                    catch (Exception ex)
                    {
                        result.FailedFiles.Add(file);
                        Console.WriteLine($"处理 {Path.GetFileName(file)} 时出错: {ex.Message}");
                    }
                }

                Console.WriteLine($"\n=== 自定义工作流执行完成 ===");
                Console.WriteLine($"成功处理: {result.ProcessedFiles.Count} 个文档");
                Console.WriteLine($"处理失败: {result.FailedFiles.Count} 个文档");
                Console.WriteLine($"总计处理: {allFiles.Length} 个文档");

                result.Success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自定义工作流执行过程中出错: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 执行自定义文档工作流
        /// </summary>
        /// <param name="inputFilePath">输入文件路径</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="customWorkflow">自定义工作流</param>
        private static void ExecuteCustomDocumentWorkflow(
            string inputFilePath,
            string outputDirectory,
            CustomWorkflow customWorkflow)
        {
            using var app = WordFactory.Open(inputFilePath);
            var document = app.ActiveDocument;

            // 执行自定义步骤
            foreach (var step in customWorkflow.Steps)
            {
                switch (step.Type)
                {
                    case WorkflowStepType.StandardizeFormat:
                        StandardizeDocumentFormat(document);
                        break;
                    case WorkflowStepType.UpdateFields:
                        UpdateDocumentFields(document);
                        break;
                    case WorkflowStepType.AddHeaderFooter:
                        AddHeaderFooter(document);
                        break;
                    case WorkflowStepType.GenerateTableOfContents:
                        GenerateTableOfContents(document);
                        break;
                    case WorkflowStepType.AddWatermark:
                        if (!string.IsNullOrEmpty(step.Parameter))
                        {
                            AddWatermark(document, step.Parameter);
                        }
                        break;
                }
            }

            // 生成输出文件路径
            var fileName = Path.GetFileNameWithoutExtension(inputFilePath);
            var outputFilePath = Path.Combine(outputDirectory, $"{fileName}_{customWorkflow.Name}.docx");

            // 保存处理后的文档
            document.SaveAs(outputFilePath);
        }

        /// <summary>
        /// 生成工作流报告
        /// </summary>
        /// <param name="result">工作流结果</param>
        /// <returns>工作流报告</returns>
        public static string GenerateWorkflowReport(WorkflowResult result)
        {
            var report = new StringBuilder();
            report.AppendLine("=== 文档自动化工作流报告 ===");
            report.AppendLine($"执行时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"总文件数: {result.TotalFiles}");
            report.AppendLine($"成功处理: {result.ProcessedFiles.Count}");
            report.AppendLine($"处理失败: {result.FailedFiles.Count}");
            report.AppendLine($"成功率: {(result.TotalFiles > 0 ? (double)result.ProcessedFiles.Count / result.TotalFiles * 100 : 0):F2}%");

            if (result.FailedFiles.Any())
            {
                report.AppendLine("\n失败文件列表:");
                foreach (var file in result.FailedFiles)
                {
                    report.AppendLine($"  - {Path.GetFileName(file)}");
                }
            }

            if (!string.IsNullOrEmpty(result.ErrorMessage))
            {
                report.AppendLine($"\n错误信息: {result.ErrorMessage}");
            }

            report.AppendLine("==========================");

            return report.ToString();
        }
    }

    /// <summary>
    /// 工作流结果类
    /// </summary>
    public class WorkflowResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 总文件数
        /// </summary>
        public int TotalFiles { get; set; }

        /// <summary>
        /// 已处理文件列表
        /// </summary>
        public List<string> ProcessedFiles { get; set; } = new List<string>();

        /// <summary>
        /// 失败文件列表
        /// </summary>
        public List<string> FailedFiles { get; set; } = new List<string>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 自定义工作流类
    /// </summary>
    public class CustomWorkflow
    {
        /// <summary>
        /// 工作流名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 工作流步骤列表
        /// </summary>
        public List<WorkflowStep> Steps { get; set; } = new List<WorkflowStep>();
    }

    /// <summary>
    /// 工作流步骤类
    /// </summary>
    public class WorkflowStep
    {
        /// <summary>
        /// 步骤类型
        /// </summary>
        public WorkflowStepType Type { get; set; }

        /// <summary>
        /// 步骤参数
        /// </summary>
        public string Parameter { get; set; }
    }

    /// <summary>
    /// 工作流步骤类型枚举
    /// </summary>
    public enum WorkflowStepType
    {
        /// <summary>
        /// 标准化格式
        /// </summary>
        StandardizeFormat,

        /// <summary>
        /// 更新字段
        /// </summary>
        UpdateFields,

        /// <summary>
        /// 添加页眉页脚
        /// </summary>
        AddHeaderFooter,

        /// <summary>
        /// 生成目录
        /// </summary>
        GenerateTableOfContents,

        /// <summary>
        /// 添加水印
        /// </summary>
        AddWatermark
    }
}