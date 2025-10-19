//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace ReportGenerationSystemSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 报表生成系统示例");

            // 示例1: 模板设计
            Console.WriteLine("\n=== 示例1: 模板设计 ===");
            TemplateDesignDemo();

            // 示例2: 数据填充
            Console.WriteLine("\n=== 示例2: 数据填充 ===");
            DataFillingDemo();

            // 示例3: 格式化处理
            Console.WriteLine("\n=== 示例3: 格式化处理 ===");
            FormattingDemo();

            // 示例4: 批量导出
            Console.WriteLine("\n=== 示例4: 批量导出 ===");
            BatchExportDemo();

            // 示例5: 实际应用示例
            Console.WriteLine("\n=== 示例5: 实际应用示例 ===");
            RealWorldApplicationDemo();

            // 示例6: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例6: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 模板设计示例
        /// </summary>
        static void TemplateDesignDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(AppContext.BaseDirectory, "ReportGenerationSystem");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                using var app = WordFactory.BlankWorkbook();
                var templateManager = new ReportTemplateManager(app);

                // 创建销售报表模板
                string salesTemplatePath = Path.Combine(tempDirectory, "SalesReportTemplate.dotx");
                bool salesTemplateCreated = templateManager.CreateSalesReportTemplate(salesTemplatePath);
                Console.WriteLine($"销售报表模板创建结果: {salesTemplateCreated}");

                // 创建财务报表模板
                string financialTemplatePath = Path.Combine(tempDirectory, "FinancialReportTemplate.dotx");
                bool financialTemplateCreated = templateManager.CreateFinancialReportTemplate(financialTemplatePath);
                Console.WriteLine($"财务报表模板创建结果: {financialTemplateCreated}");

                // 创建项目进度报表模板
                string projectTemplatePath = Path.Combine(tempDirectory, "ProjectProgressReportTemplate.dotx");
                bool projectTemplateCreated = templateManager.CreateProjectProgressReportTemplate(projectTemplatePath);
                Console.WriteLine($"项目进度报表模板创建结果: {projectTemplateCreated}");

                // 获取模板信息
                var templateInfo = templateManager.GetTemplateInfo(salesTemplatePath);
                Console.WriteLine(templateInfo.GenerateReport());

                Console.WriteLine("模板设计示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"模板设计示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 数据填充示例
        /// </summary>
        static void DataFillingDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(AppContext.BaseDirectory, "ReportGenerationSystem");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 确保模板存在
                string templatePath = Path.Combine(tempDirectory, "SalesReportTemplate.dotx");
                if (!File.Exists(templatePath))
                {
                    using var app = WordFactory.BlankWorkbook();
                    var templateManager = new ReportTemplateManager(app);
                    templateManager.CreateSalesReportTemplate(templatePath);
                }

                var dataFiller = new ReportDataFiller();

                // 生成销售报表
                string outputPath = Path.Combine(tempDirectory, "SampleSalesReport.docx");
                bool reportGenerated = dataFiller.GenerateSalesReport(templatePath, outputPath, DateTime.Now.AddMonths(-1));
                Console.WriteLine($"销售报表生成结果: {reportGenerated}");

                // 生成财务报表
                string financialTemplatePath = Path.Combine(tempDirectory, "FinancialReportTemplate.dotx");
                if (!File.Exists(financialTemplatePath))
                {
                    using var app = WordFactory.BlankWorkbook();
                    var templateManager = new ReportTemplateManager(app);
                    templateManager.CreateFinancialReportTemplate(financialTemplatePath);
                }

                string financialOutputPath = Path.Combine(tempDirectory, "SampleFinancialReport.docx");
                bool financialReportGenerated = dataFiller.GenerateFinancialReport(financialTemplatePath, financialOutputPath, DateTime.Now.AddMonths(-1));
                Console.WriteLine($"财务报表生成结果: {financialReportGenerated}");

                // 生成项目进度报表
                string projectTemplatePath = Path.Combine(tempDirectory, "ProjectProgressReportTemplate.dotx");
                if (!File.Exists(projectTemplatePath))
                {
                    using var app = WordFactory.BlankWorkbook();
                    var templateManager = new ReportTemplateManager(app);
                    templateManager.CreateProjectProgressReportTemplate(projectTemplatePath);
                }

                string projectOutputPath = Path.Combine(tempDirectory, "SampleProjectProgressReport.docx");
                bool projectReportGenerated = dataFiller.GenerateProjectProgressReport(projectTemplatePath, projectOutputPath, "ERP系统开发项目", DateTime.Now);
                Console.WriteLine($"项目进度报表生成结果: {projectReportGenerated}");

                Console.WriteLine("数据填充示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"数据填充示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 格式化处理示例
        /// </summary>
        static void FormattingDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(AppContext.BaseDirectory, "ReportGenerationSystem");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 确保报表存在
                string reportPath = Path.Combine(tempDirectory, "SampleSalesReport.docx");
                if (!File.Exists(reportPath))
                {
                    string templatePath = Path.Combine(tempDirectory, "SalesReportTemplate.dotx");
                    if (!File.Exists(templatePath))
                    {
                        using var appb = WordFactory.BlankWorkbook();
                        var templateManager = new ReportTemplateManager(appb);
                        templateManager.CreateSalesReportTemplate(templatePath);
                    }

                    var dataFiller = new ReportDataFiller();
                    dataFiller.GenerateSalesReport(templatePath, reportPath, DateTime.Now.AddMonths(-1));
                }

                using var app = WordFactory.Open(reportPath);
                var document = app.ActiveDocument;
                var formatter = new ReportFormatter();

                // 应用专业格式化
                bool professionalFormattingApplied = formatter.ApplyProfessionalFormatting(document);
                Console.WriteLine($"专业格式化应用结果: {professionalFormattingApplied}");

                // 应用现代风格格式化
                bool modernFormattingApplied = formatter.ApplyModernFormatting(document);
                Console.WriteLine($"现代风格格式化应用结果: {modernFormattingApplied}");

                // 应用简洁风格格式化
                bool minimalistFormattingApplied = formatter.ApplyMinimalistFormatting(document);
                Console.WriteLine($"简洁风格格式化应用结果: {minimalistFormattingApplied}");

                // 应用品牌格式化
                bool brandingFormattingApplied = formatter.ApplyBrandingFormatting(document, "ABC公司", WdColor.wdColorDarkBlue);
                Console.WriteLine($"品牌格式化应用结果: {brandingFormattingApplied}");

                // 保存格式化后的报表
                string formattedReportPath = Path.Combine(tempDirectory, "FormattedSalesReport.docx");
                document.SaveAs(formattedReportPath);
                Console.WriteLine($"格式化后的报表已保存: {formattedReportPath}");

                Console.WriteLine("格式化处理示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式化处理示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 批量导出示例
        /// </summary>
        static void BatchExportDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(AppContext.BaseDirectory, "ReportGenerationSystem");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                string batchOutputDirectory = Path.Combine(tempDirectory, "BatchReports");
                if (!Directory.Exists(batchOutputDirectory))
                {
                    Directory.CreateDirectory(batchOutputDirectory);
                }

                // 确保模板存在
                string templatePath = Path.Combine(tempDirectory, "SalesReportTemplate.dotx");
                if (!File.Exists(templatePath))
                {
                    using var app = WordFactory.BlankWorkbook();
                    var templateManager1 = new ReportTemplateManager(app);
                    templateManager1.CreateSalesReportTemplate(templatePath);
                }

                var templateManager = new ReportTemplateManager(null);
                var dataFiller = new ReportDataFiller();
                var formatter = new ReportFormatter();
                var batchGenerator = new BatchReportGenerator(templateManager, dataFiller, formatter);

                // 生成年度销售报表
                string annualReportsDirectory = Path.Combine(batchOutputDirectory, "AnnualSalesReports");
                bool annualReportsGenerated = batchGenerator.GenerateAnnualSalesReports(DateTime.Now.Year, templatePath, annualReportsDirectory);
                Console.WriteLine($"年度销售报表生成结果: {annualReportsGenerated}");

                // 生成部门销售报表
                string departmentReportsDirectory = Path.Combine(batchOutputDirectory, "DepartmentSalesReports");
                bool departmentReportsGenerated = batchGenerator.GenerateDepartmentSalesReports(templatePath, DateTime.Now, departmentReportsDirectory);
                Console.WriteLine($"部门销售报表生成结果: {departmentReportsGenerated}");

                // 生成项目进度报表
                string projectReportsDirectory = Path.Combine(batchOutputDirectory, "ProjectProgressReports");
                var projects = new List<string> { "ERP系统", "CRM系统", "OA系统", "电商平台", "移动应用" };
                bool projectReportsGenerated = batchGenerator.GenerateProjectProgressReports(
                    Path.Combine(tempDirectory, "ProjectProgressReportTemplate.dotx"),
                    projects,
                    DateTime.Now,
                    projectReportsDirectory);
                Console.WriteLine($"项目进度报表生成结果: {projectReportsGenerated}");

                // 生成季度财务报表
                string quarterlyReportsDirectory = Path.Combine(batchOutputDirectory, "QuarterlyFinancialReports");
                var quarters = new List<string> { "一", "二", "三", "四" };
                bool quarterlyReportsGenerated = batchGenerator.GenerateQuarterlyFinancialReports(
                    Path.Combine(tempDirectory, "FinancialReportTemplate.dotx"),
                    DateTime.Now,
                    quarterlyReportsDirectory,
                    quarters);
                Console.WriteLine($"季度财务报表生成结果: {quarterlyReportsGenerated}");

                // 生成进度报告
                string progressReport = batchGenerator.GenerateProgressReport(10, 12);
                Console.WriteLine(progressReport);

                // 生成总结报告
                string summaryReport = batchGenerator.GenerateSummaryReport(10, 2, 12, batchOutputDirectory);
                Console.WriteLine(summaryReport);

                Console.WriteLine("批量导出示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量导出示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例
        /// </summary>
        static void RealWorldApplicationDemo()
        {
            try
            {
                Console.WriteLine("=== 报表生成系统演示 ===");
                Console.WriteLine();

                // 创建临时目录
                string tempDirectory = Path.Combine(AppContext.BaseDirectory, "ReportGenerationSystem");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 步骤1: 创建模板
                Console.WriteLine("步骤1: 创建报表模板");
                using var app = WordFactory.BlankWorkbook();
                var templateManager = new ReportTemplateManager(app);
                string templatePath = Path.Combine(tempDirectory, "SalesReportTemplate.dotx");
                templateManager.CreateSalesReportTemplate(templatePath);
                Console.WriteLine();

                // 步骤2: 生成单个报表
                Console.WriteLine("步骤2: 生成单个报表");
                var dataFiller = new ReportDataFiller();
                string outputPath = Path.Combine(tempDirectory, "SampleSalesReport.docx");
                dataFiller.GenerateSalesReport(templatePath, outputPath, DateTime.Now.AddMonths(-1));
                Console.WriteLine();

                // 步骤3: 应用专业格式化
                Console.WriteLine("步骤3: 应用专业格式化");
                using var formattedApp = WordFactory.Open(outputPath);
                var document = formattedApp.ActiveDocument;
                var formatter = new ReportFormatter();
                formatter.ApplyProfessionalFormatting(document);
                document.Save();
                Console.WriteLine();

                // 步骤4: 批量生成报表
                Console.WriteLine("步骤4: 批量生成报表");
                var batchGenerator = new BatchReportGenerator(templateManager, dataFiller, formatter);
                string batchOutputDirectory = Path.Combine(tempDirectory, "BatchReports");
                batchGenerator.GenerateAnnualSalesReports(DateTime.Now.Year, templatePath, batchOutputDirectory);
                Console.WriteLine();

                Console.WriteLine("报表生成系统演示完成！");
                Console.WriteLine($"生成的报表位于 {tempDirectory} 目录下");
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
                string tempDirectory = Path.Combine(AppContext.BaseDirectory, "ReportGenerationSystem");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 创建各个管理器
                using var app = WordFactory.BlankWorkbook();
                var templateManager = new ReportTemplateManager(app);
                var dataFiller = new ReportDataFiller();
                var formatter = new ReportFormatter();

                // 1. 创建多种模板
                Console.WriteLine("1. 创建多种报表模板");
                string salesTemplatePath = Path.Combine(tempDirectory, "SalesReportTemplate.dotx");
                templateManager.CreateSalesReportTemplate(salesTemplatePath);

                string financialTemplatePath = Path.Combine(tempDirectory, "FinancialReportTemplate.dotx");
                templateManager.CreateFinancialReportTemplate(financialTemplatePath);

                string projectTemplatePath = Path.Combine(tempDirectory, "ProjectProgressReportTemplate.dotx");
                templateManager.CreateProjectProgressReportTemplate(projectTemplatePath);

                // 2. 获取模板信息
                Console.WriteLine("2. 获取模板信息");
                var templateInfo = templateManager.GetTemplateInfo(salesTemplatePath);
                Console.WriteLine(templateInfo.GenerateReport());

                // 3. 生成各种类型的报表
                Console.WriteLine("3. 生成各种类型的报表");
                string salesReportPath = Path.Combine(tempDirectory, "CompleteSalesReport.docx");
                dataFiller.GenerateSalesReport(salesTemplatePath, salesReportPath, DateTime.Now.AddMonths(-1));

                string financialReportPath = Path.Combine(tempDirectory, "CompleteFinancialReport.docx");
                dataFiller.GenerateFinancialReport(financialTemplatePath, financialReportPath, DateTime.Now.AddMonths(-1));

                string projectReportPath = Path.Combine(tempDirectory, "CompleteProjectReport.docx");
                dataFiller.GenerateProjectProgressReport(projectTemplatePath, projectReportPath, "完整项目", DateTime.Now);

                // 4. 应用格式化
                Console.WriteLine("4. 应用格式化");
                using var formattedApp = WordFactory.Open(salesReportPath);
                var document = formattedApp.ActiveDocument;
                formatter.ApplyProfessionalFormatting(document);
                document.Save();

                // 5. 批量生成
                Console.WriteLine("5. 批量生成报表");
                var batchGenerator = new BatchReportGenerator(templateManager, dataFiller, formatter);
                string batchOutputDirectory = Path.Combine(tempDirectory, "CompleteBatchReports");
                batchGenerator.GenerateDepartmentSalesReports(salesTemplatePath, DateTime.Now, batchOutputDirectory);

                // 6. 生成报告
                Console.WriteLine("6. 生成进度和总结报告");
                string progressReport = batchGenerator.GenerateProgressReport(5, 5);
                Console.WriteLine(progressReport);

                string summaryReport = batchGenerator.GenerateSummaryReport(5, 0, 5, batchOutputDirectory);
                Console.WriteLine(summaryReport);

                Console.WriteLine("使用辅助类的完整示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例演示出错: {ex.Message}");
            }
        }
    }
}