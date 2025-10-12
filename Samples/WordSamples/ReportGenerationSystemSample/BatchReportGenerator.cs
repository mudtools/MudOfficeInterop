using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerationSystemSample
{
    /// <summary>
    /// 批量报表生成器类
    /// </summary>
    public class BatchReportGenerator
    {
        private readonly ReportTemplateManager _templateManager;
        private readonly ReportDataFiller _dataFiller;
        private readonly ReportFormatter _formatter;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="templateManager">报表模板管理器</param>
        /// <param name="dataFiller">报表数据填充器</param>
        /// <param name="formatter">报表格式化器</param>
        public BatchReportGenerator(ReportTemplateManager templateManager, ReportDataFiller dataFiller, ReportFormatter formatter)
        {
            _templateManager = templateManager ?? throw new ArgumentNullException(nameof(templateManager));
            _dataFiller = dataFiller ?? throw new ArgumentNullException(nameof(dataFiller));
            _formatter = formatter ?? throw new ArgumentNullException(nameof(formatter));
        }

        /// <summary>
        /// 生成年度销售报表
        /// </summary>
        /// <param name="year">年份</param>
        /// <param name="templatePath">模板路径</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>是否生成成功</returns>
        public bool GenerateAnnualSalesReports(int year, string templatePath, string outputDirectory)
        {
            // 确保输出目录存在
            if (!System.IO.Directory.Exists(outputDirectory))
            {
                System.IO.Directory.CreateDirectory(outputDirectory);
            }

            try
            {
                Console.WriteLine($"开始生成 {year} 年度销售报表...");

                // 为每个月生成报表
                for (int month = 1; month <= 12; month++)
                {
                    var reportPeriod = new DateTime(year, month, 1);
                    string outputPath = System.IO.Path.Combine(outputDirectory, $"{year}年{month}月销售报表.docx");

                    Console.WriteLine($"  正在生成 {year}年{month}月 销售报表...");

                    // 生成单个报表
                    bool success = _dataFiller.GenerateSalesReport(templatePath, outputPath, reportPeriod);
                    if (success)
                    {
                        Console.WriteLine($"  已完成 {year}年{month}月 销售报表");
                    }
                    else
                    {
                        Console.WriteLine($"  生成 {year}年{month}月 销售报表失败");
                    }
                }

                Console.WriteLine($"所有 {year} 年度销售报表已生成到: {outputDirectory}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量生成年度销售报表时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成部门销售报表
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="reportPeriod">报表期间</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>是否生成成功</returns>
        public bool GenerateDepartmentSalesReports(string templatePath, DateTime reportPeriod, string outputDirectory)
        {
            // 确保输出目录存在
            if (!System.IO.Directory.Exists(outputDirectory))
            {
                System.IO.Directory.CreateDirectory(outputDirectory);
            }

            // 部门列表
            var departments = new List<(string Name, string Code)>
            {
                ("销售部", "SALES"),
                ("市场部", "MARKETING"),
                ("技术部", "TECH"),
                ("人事部", "HR"),
                ("财务部", "FINANCE")
            };

            try
            {
                Console.WriteLine($"开始生成 {reportPeriod:yyyy年MM月} 部门销售报表...");

                foreach (var department in departments)
                {
                    string outputPath = System.IO.Path.Combine(outputDirectory, $"{department.Name}销售报表.docx");

                    Console.WriteLine($"  正在生成 {department.Name} 销售报表...");

                    using var app = WordFactory.CreateFrom(templatePath);
                    var document = app.ActiveDocument;

                    // 自定义每个部门的报表内容
                    CustomizeDepartmentSalesReport(document, department.Name, department.Code, reportPeriod);

                    // 应用格式化
                    _formatter.ApplyProfessionalFormatting(document);

                    // 保存报表
                    document.SaveAs2(outputPath);

                    Console.WriteLine($"  {department.Name} 销售报表已生成: {outputPath}");
                }

                Console.WriteLine($"所有部门销售报表已生成到: {outputDirectory}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成部门销售报表时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 自定义部门销售报表内容
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="departmentName">部门名称</param>
        /// <param name="departmentCode">部门代码</param>
        /// <param name="reportPeriod">报表期间</param>
        private void CustomizeDepartmentSalesReport(IWordDocument document, string departmentName, string departmentCode, DateTime reportPeriod)
        {
            try
            {
                // 替换部门特定内容
                var range = document.Range();
                var text = range.Text;

                text = text.Replace("{DEPARTMENT_NAME}", departmentName);
                text = text.Replace("{DEPARTMENT_CODE}", departmentCode);
                text = text.Replace("{REPORT_PERIOD}", $"{reportPeriod:yyyy年MM月}");

                range.Text = text;

                // 可以根据部门添加特定内容
                switch (departmentName)
                {
                    case "销售部":
                        AddSalesDepartmentSpecificContent(document);
                        break;
                    case "市场部":
                        AddMarketingDepartmentSpecificContent(document);
                        break;
                    case "技术部":
                        AddTechDepartmentSpecificContent(document);
                        break;
                    case "人事部":
                        AddHRDepartmentSpecificContent(document);
                        break;
                    case "财务部":
                        AddFinanceDepartmentSpecificContent(document);
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自定义部门销售报表内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加销售部特定内容
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void AddSalesDepartmentSpecificContent(IWordDocument document)
        {
            try
            {
                var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                range.Text = "\n\n销售业绩分析:\n" +
                            "• 本月销售额达到预期目标\n" +
                            "• 新客户开发数量同比增长15%\n" +
                            "• 客户满意度保持在95%以上\n";
                range.ListFormat.ApplyBulletDefault();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加销售部特定内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加市场部特定内容
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void AddMarketingDepartmentSpecificContent(IWordDocument document)
        {
            try
            {
                var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                range.Text = "\n\n市场活动分析:\n" +
                            "• 本月成功举办3场市场推广活动\n" +
                            "• 品牌知名度提升10%\n" +
                            "• 市场占有率稳定在预期水平\n";
                range.ListFormat.ApplyBulletDefault();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加市场部特定内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加技术部特定内容
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void AddTechDepartmentSpecificContent(IWordDocument document)
        {
            try
            {
                var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                range.Text = "\n\n技术研发分析:\n" +
                            "• 本月完成2个重要功能开发\n" +
                            "• 系统稳定性提升，故障率下降20%\n" +
                            "• 技术债务减少15%\n";
                range.ListFormat.ApplyBulletDefault();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加技术部特定内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加人事部特定内容
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void AddHRDepartmentSpecificContent(IWordDocument document)
        {
            try
            {
                var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                range.Text = "\n\n人力资源分析:\n" +
                            "• 本月新招聘员工8人\n" +
                            "• 员工满意度调查得分92分\n" +
                            "• 培训完成率达到95%\n";
                range.ListFormat.ApplyBulletDefault();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加人事部特定内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加财务部特定内容
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void AddFinanceDepartmentSpecificContent(IWordDocument document)
        {
            try
            {
                var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                range.Text = "\n\n财务状况分析:\n" +
                            "• 现金流状况良好\n" +
                            "• 成本控制效果显著\n" +
                            "• 投资回报率稳步提升\n";
                range.ListFormat.ApplyBulletDefault();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加财务部特定内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 生成项目进度报表
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="projects">项目列表</param>
        /// <param name="reportPeriod">报表期间</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>是否生成成功</returns>
        public bool GenerateProjectProgressReports(string templatePath, List<string> projects, DateTime reportPeriod, string outputDirectory)
        {
            // 确保输出目录存在
            if (!System.IO.Directory.Exists(outputDirectory))
            {
                System.IO.Directory.CreateDirectory(outputDirectory);
            }

            try
            {
                Console.WriteLine($"开始生成项目进度报表...");

                foreach (var project in projects)
                {
                    string outputPath = System.IO.Path.Combine(outputDirectory, $"{project}进度报表.docx");

                    Console.WriteLine($"  正在生成 {project} 进度报表...");

                    // 生成单个项目进度报表
                    bool success = _dataFiller.GenerateProjectProgressReport(templatePath, outputPath, project, reportPeriod);
                    if (success)
                    {
                        Console.WriteLine($"  {project} 进度报表已生成: {outputPath}");
                    }
                    else
                    {
                        Console.WriteLine($"  生成 {project} 进度报表失败");
                    }
                }

                Console.WriteLine($"所有项目进度报表已生成到: {outputDirectory}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成项目进度报表时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成财务报表
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="reportPeriod">报表期间</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="quarters">季度列表</param>
        /// <returns>是否生成成功</returns>
        public bool GenerateQuarterlyFinancialReports(string templatePath, DateTime reportPeriod, string outputDirectory, List<string> quarters)
        {
            // 确保输出目录存在
            if (!System.IO.Directory.Exists(outputDirectory))
            {
                System.IO.Directory.CreateDirectory(outputDirectory);
            }

            try
            {
                Console.WriteLine($"开始生成季度财务报表...");

                foreach (var quarter in quarters)
                {
                    string outputPath = System.IO.Path.Combine(outputDirectory, $"{reportPeriod.Year}年第{quarter}季度财务报表.docx");

                    Console.WriteLine($"  正在生成 {reportPeriod.Year}年第{quarter}季度 财务报表...");

                    using var app = WordFactory.CreateFrom(templatePath);
                    var document = app.ActiveDocument;

                    // 自定义每个季度的报表内容
                    CustomizeQuarterlyFinancialReport(document, quarter, reportPeriod);

                    // 应用格式化
                    _formatter.ApplyProfessionalFormatting(document);

                    // 保存报表
                    document.SaveAs2(outputPath);

                    Console.WriteLine($"  {reportPeriod.Year}年第{quarter}季度 财务报表已生成: {outputPath}");
                }

                Console.WriteLine($"所有季度财务报表已生成到: {outputDirectory}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成季度财务报表时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 自定义季度财务报表内容
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="quarter">季度</param>
        /// <param name="reportPeriod">报表期间</param>
        private void CustomizeQuarterlyFinancialReport(IWordDocument document, string quarter, DateTime reportPeriod)
        {
            try
            {
                // 替换季度特定内容
                var range = document.Range();
                var text = range.Text;

                text = text.Replace("{QUARTER}", quarter);
                text = text.Replace("{REPORT_PERIOD}", $"{reportPeriod.Year}年第{quarter}季度");

                range.Text = text;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自定义季度财务报表内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 生成批量报表的进度报告
        /// </summary>
        /// <param name="reportsGenerated">已生成报表数量</param>
        /// <param name="totalReports">总报表数量</param>
        /// <returns>进度报告</returns>
        public string GenerateProgressReport(int reportsGenerated, int totalReports)
        {
            double progress = totalReports > 0 ? (double)reportsGenerated / totalReports * 100 : 0;
            return $"报表生成进度: {reportsGenerated}/{totalReports} ({progress:F1}%)";
        }

        /// <summary>
        /// 生成批量报表的总结报告
        /// </summary>
        /// <param name="successfulReports">成功生成的报表数量</param>
        /// <param name="failedReports">生成失败的报表数量</param>
        /// <param name="totalReports">总报表数量</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>总结报告</returns>
        public string GenerateSummaryReport(int successfulReports, int failedReports, int totalReports, string outputDirectory)
        {
            double successRate = totalReports > 0 ? (double)successfulReports / totalReports * 100 : 0;
            return $"报表生成总结报告:\n" +
                   $"  总报表数量: {totalReports}\n" +
                   $"  成功生成: {successfulReports}\n" +
                   $"  生成失败: {failedReports}\n" +
                   $"  成功率: {successRate:F1}%\n" +
                   $"  输出目录: {outputDirectory}";
        }
    }
}