using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerationSystemSample
{
    /// <summary>
    /// 报表数据填充器类
    /// </summary>
    public class ReportDataFiller
    {
        /// <summary>
        /// 销售数据类
        /// </summary>
        public class SalesData
        {
            /// <summary>
            /// 产品名称
            /// </summary>
            public string ProductName { get; set; }

            /// <summary>
            /// 销售数量
            /// </summary>
            public int Quantity { get; set; }

            /// <summary>
            /// 单价
            /// </summary>
            public decimal UnitPrice { get; set; }

            /// <summary>
            /// 总金额
            /// </summary>
            public decimal TotalAmount { get; set; }

            /// <summary>
            /// 增长率
            /// </summary>
            public decimal GrowthRate { get; set; }
        }

        /// <summary>
        /// 财务数据类
        /// </summary>
        public class FinancialData
        {
            /// <summary>
            /// 收入项目
            /// </summary>
            public string IncomeItem { get; set; }

            /// <summary>
            /// 收入金额
            /// </summary>
            public decimal IncomeAmount { get; set; }

            /// <summary>
            /// 支出项目
            /// </summary>
            public string ExpenseItem { get; set; }

            /// <summary>
            /// 支出金额
            /// </summary>
            public decimal ExpenseAmount { get; set; }
        }

        /// <summary>
        /// 项目进度数据类
        /// </summary>
        public class ProjectProgressData
        {
            /// <summary>
            /// 任务名称
            /// </summary>
            public string TaskName { get; set; }

            /// <summary>
            /// 负责人
            /// </summary>
            public string Assignee { get; set; }

            /// <summary>
            /// 计划开始时间
            /// </summary>
            public DateTime PlannedStart { get; set; }

            /// <summary>
            /// 实际开始时间
            /// </summary>
            public DateTime ActualStart { get; set; }

            /// <summary>
            /// 计划完成时间
            /// </summary>
            public DateTime PlannedEnd { get; set; }

            /// <summary>
            /// 预计完成时间
            /// </summary>
            public DateTime EstimatedEnd { get; set; }

            /// <summary>
            /// 完成百分比
            /// </summary>
            public int CompletionPercentage { get; set; }

            /// <summary>
            /// 状态
            /// </summary>
            public string Status { get; set; }
        }

        /// <summary>
        /// 生成销售报表
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="outputPath">输出路径</param>
        /// <param name="reportPeriod">报表期间</param>
        /// <returns>是否生成成功</returns>
        public bool GenerateSalesReport(string templatePath, string outputPath, DateTime reportPeriod)
        {
            try
            {
                using var app = WordFactory.CreateFrom(templatePath);
                var document = app.ActiveDocument;

                // 填充报表信息
                FillReportInfo(document, reportPeriod, "月度销售报表");

                // 填充销售数据
                var salesData = GenerateSampleSalesData();
                FillSalesData(document, salesData);

                // 填充总结信息
                FillSalesSummaryInfo(document, salesData);

                // 保存报表
                document.SaveAs2(outputPath);

                Console.WriteLine($"销售报表已生成: {outputPath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成销售报表时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成财务报表
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="outputPath">输出路径</param>
        /// <param name="reportPeriod">报表期间</param>
        /// <returns>是否生成成功</returns>
        public bool GenerateFinancialReport(string templatePath, string outputPath, DateTime reportPeriod)
        {
            try
            {
                using var app = WordFactory.CreateFrom(templatePath);
                var document = app.ActiveDocument;

                // 填充报表信息
                FillReportInfo(document, reportPeriod, "月度财务报表");

                // 填充财务数据
                var financialData = GenerateSampleFinancialData();
                FillFinancialData(document, financialData);

                // 填充总结信息
                FillFinancialSummaryInfo(document, financialData);

                // 保存报表
                document.SaveAs2(outputPath);

                Console.WriteLine($"财务报表已生成: {outputPath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成财务报表时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成项目进度报表
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="outputPath">输出路径</param>
        /// <param name="projectName">项目名称</param>
        /// <param name="reportPeriod">报表期间</param>
        /// <returns>是否生成成功</returns>
        public bool GenerateProjectProgressReport(string templatePath, string outputPath, string projectName, DateTime reportPeriod)
        {
            try
            {
                using var app = WordFactory.CreateFrom(templatePath);
                var document = app.ActiveDocument;

                // 填充报表信息
                FillProjectReportInfo(document, projectName, reportPeriod);

                // 填充项目数据
                var projectData = GenerateSampleProjectProgressData();
                FillProjectProgressData(document, projectData);

                // 填充里程碑数据
                FillMilestones(document);

                // 填充风险和问题
                FillRisksAndIssues(document);

                // 填充总结信息
                FillProjectSummaryInfo(document, projectData);

                // 保存报表
                document.SaveAs2(outputPath);

                Console.WriteLine($"项目进度报表已生成: {outputPath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成项目进度报表时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 填充报表基本信息
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="reportPeriod">报表期间</param>
        /// <param name="reportType">报表类型</param>
        private void FillReportInfo(IWordDocument document, DateTime reportPeriod, string reportType)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                // 替换占位符
                text = text.Replace("{REPORT_PERIOD}", $"{reportPeriod:yyyy年MM月}");
                text = text.Replace("{GENERATION_TIME}", $"{DateTime.Now:yyyy年MM月dd日 HH:mm:ss}");
                text = text.Replace("{REPORT_TYPE}", reportType);

                range.Text = text;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充报表信息时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充项目报表基本信息
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="projectName">项目名称</param>
        /// <param name="reportPeriod">报表期间</param>
        private void FillProjectReportInfo(IWordDocument document, string projectName, DateTime reportPeriod)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                // 替换占位符
                text = text.Replace("{PROJECT_NAME}", projectName);
                text = text.Replace("{REPORT_PERIOD}", $"{reportPeriod:yyyy年MM月}");
                text = text.Replace("{GENERATION_TIME}", $"{DateTime.Now:yyyy年MM月dd日 HH:mm:ss}");

                range.Text = text;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充项目报表信息时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 生成示例销售数据
        /// </summary>
        /// <returns>销售数据列表</returns>
        private List<SalesData> GenerateSampleSalesData()
        {
            return new List<SalesData>
            {
                new SalesData { ProductName = "产品A", Quantity = 1000, UnitPrice = 50.00m, TotalAmount = 50000.00m, GrowthRate = 0.15m },
                new SalesData { ProductName = "产品B", Quantity = 800, UnitPrice = 75.00m, TotalAmount = 60000.00m, GrowthRate = 0.12m },
                new SalesData { ProductName = "产品C", Quantity = 1200, UnitPrice = 40.00m, TotalAmount = 48000.00m, GrowthRate = 0.08m },
                new SalesData { ProductName = "产品D", Quantity = 600, UnitPrice = 100.00m, TotalAmount = 60000.00m, GrowthRate = 0.20m },
                new SalesData { ProductName = "产品E", Quantity = 1500, UnitPrice = 30.00m, TotalAmount = 45000.00m, GrowthRate = 0.05m }
            };
        }

        /// <summary>
        /// 填充销售数据
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="salesData">销售数据</param>
        private void FillSalesData(IWordDocument document, List<SalesData> salesData)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                // 查找表格占位符位置
                int tablePosition = text.IndexOf("{SALES_DATA_TABLE}");
                if (tablePosition >= 0)
                {
                    // 创建表格
                    var tableRange = document.Range(tablePosition, tablePosition + 18); // 18是"{SALES_DATA_TABLE}"的长度
                    var table = document.Tables.Add(tableRange, salesData.Count + 1, 5); // 表头+数据行

                    // 设置表头
                    table.Cell(1, 1).Range.Text = "产品名称";
                    table.Cell(1, 2).Range.Text = "销售数量";
                    table.Cell(1, 3).Range.Text = "单价(元)";
                    table.Cell(1, 4).Range.Text = "总金额(元)";
                    table.Cell(1, 5).Range.Text = "增长率";

                    // 格式化表头
                    for (int i = 1; i <= 5; i++)
                    {
                        var cell = table.Cell(1, i);
                        cell.Range.Font.Bold = 1;
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                    }

                    // 填充数据
                    for (int i = 0; i < salesData.Count; i++)
                    {
                        var data = salesData[i];
                        table.Cell(i + 2, 1).Range.Text = data.ProductName;
                        table.Cell(i + 2, 2).Range.Text = data.Quantity.ToString();
                        table.Cell(i + 2, 3).Range.Text = data.UnitPrice.ToString("F2");
                        table.Cell(i + 2, 4).Range.Text = data.TotalAmount.ToString("F2");
                        table.Cell(i + 2, 5).Range.Text = $"{data.GrowthRate:P2}";

                        // 格式化数据行
                        for (int j = 1; j <= 5; j++)
                        {
                            table.Cell(i + 2, j).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                    }

                    // 设置表格样式
                    table.Borders.Enable = 1;
                    table.AllowAutoFit = true;
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充销售数据时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充销售总结信息
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="salesData">销售数据</param>
        private void FillSalesSummaryInfo(IWordDocument document, List<SalesData> salesData)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                // 计算总结数据
                decimal totalSales = salesData.Sum(d => d.TotalAmount);
                decimal avgGrowth = salesData.Average(d => d.GrowthRate);

                // 替换占位符
                text = text.Replace("{TOTAL_SALES}", $"{totalSales:F2} 元");
                text = text.Replace("{YEAR_OVER_YEAR_GROWTH}", $"{avgGrowth:P2}");
                text = text.Replace("{MONTH_OVER_MONTH_GROWTH}", "待计算");

                range.Text = text;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充销售总结信息时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 生成示例财务数据
        /// </summary>
        /// <returns>财务数据列表</returns>
        private List<FinancialData> GenerateSampleFinancialData()
        {
            return new List<FinancialData>
            {
                new FinancialData { IncomeItem = "销售收入", IncomeAmount = 200000.00m, ExpenseItem = "原材料成本", ExpenseAmount = 80000.00m },
                new FinancialData { IncomeItem = "服务收入", IncomeAmount = 50000.00m, ExpenseItem = "人工成本", ExpenseAmount = 60000.00m },
                new FinancialData { IncomeItem = "投资收益", IncomeAmount = 10000.00m, ExpenseItem = "管理费用", ExpenseAmount = 20000.00m },
                new FinancialData { IncomeItem = "其他收入", IncomeAmount = 5000.00m, ExpenseItem = "营销费用", ExpenseAmount = 15000.00m },
                new FinancialData { IncomeItem = "", IncomeAmount = 0m, ExpenseItem = "税费", ExpenseAmount = 12000.00m }
            };
        }

        /// <summary>
        /// 填充财务数据
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="financialData">财务数据</param>
        private void FillFinancialData(IWordDocument document, List<FinancialData> financialData)
        {
            try
            {
                // 填充收入表格
                FillIncomeTable(document, financialData);

                // 填充支出表格
                FillExpenseTable(document, financialData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充财务数据时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充收入表格
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="financialData">财务数据</param>
        private void FillIncomeTable(IWordDocument document, List<FinancialData> financialData)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                int tablePosition = text.IndexOf("{INCOME_TABLE}");
                if (tablePosition >= 0)
                {
                    var tableRange = document.Range(tablePosition, tablePosition + 14);
                    var table = document.Tables.Add(tableRange, financialData.Count + 1, 2);

                    // 设置表头
                    table.Cell(1, 1).Range.Text = "收入项目";
                    table.Cell(1, 2).Range.Text = "金额(元)";

                    // 格式化表头
                    for (int i = 1; i <= 2; i++)
                    {
                        var cell = table.Cell(1, i);
                        cell.Range.Font.Bold = 1;
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                    }

                    // 填充数据
                    decimal totalIncome = 0;
                    for (int i = 0; i < financialData.Count; i++)
                    {
                        var data = financialData[i];
                        table.Cell(i + 2, 1).Range.Text = data.IncomeItem;
                        table.Cell(i + 2, 2).Range.Text = data.IncomeAmount.ToString("F2");
                        totalIncome += data.IncomeAmount;

                        // 格式化数据行
                        table.Cell(i + 2, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        table.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    }

                    // 添加总计行
                    var totalRow = table.Rows.Add();
                    totalRow.Cells[1].Range.Text = "总收入";
                    totalRow.Cells[2].Range.Text = totalIncome.ToString("F2");
                    totalRow.Cells[1].Range.Font.Bold = 1;
                    totalRow.Cells[2].Range.Font.Bold = 1;
                    totalRow.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    totalRow.Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                    // 设置表格样式
                    table.Borders.Enable = 1;
                    table.AllowAutoFit = true;
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充收入表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充支出表格
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="financialData">财务数据</param>
        private void FillExpenseTable(IWordDocument document, List<FinancialData> financialData)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                int tablePosition = text.IndexOf("{EXPENSE_TABLE}");
                if (tablePosition >= 0)
                {
                    var tableRange = document.Range(tablePosition, tablePosition + 15);
                    var table = document.Tables.Add(tableRange, financialData.Count + 1, 2);

                    // 设置表头
                    table.Cell(1, 1).Range.Text = "支出项目";
                    table.Cell(1, 2).Range.Text = "金额(元)";

                    // 格式化表头
                    for (int i = 1; i <= 2; i++)
                    {
                        var cell = table.Cell(1, i);
                        cell.Range.Font.Bold = 1;
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                    }

                    // 填充数据
                    decimal totalExpense = 0;
                    for (int i = 0; i < financialData.Count; i++)
                    {
                        var data = financialData[i];
                        table.Cell(i + 2, 1).Range.Text = data.ExpenseItem;
                        table.Cell(i + 2, 2).Range.Text = data.ExpenseAmount.ToString("F2");
                        totalExpense += data.ExpenseAmount;

                        // 格式化数据行
                        table.Cell(i + 2, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        table.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    }

                    // 添加总计行
                    var totalRow = table.Rows.Add();
                    totalRow.Cells[1].Range.Text = "总支出";
                    totalRow.Cells[2].Range.Text = totalExpense.ToString("F2");
                    totalRow.Cells[1].Range.Font.Bold = 1;
                    totalRow.Cells[2].Range.Font.Bold = 1;
                    totalRow.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    totalRow.Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                    // 设置表格样式
                    table.Borders.Enable = 1;
                    table.AllowAutoFit = true;
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充支出表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充财务总结信息
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="financialData">财务数据</param>
        private void FillFinancialSummaryInfo(IWordDocument document, List<FinancialData> financialData)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                // 计算总结数据
                decimal totalIncome = financialData.Sum(d => d.IncomeAmount);
                decimal totalExpense = financialData.Sum(d => d.ExpenseAmount);
                decimal netProfit = totalIncome - totalExpense;
                decimal profitMargin = totalIncome > 0 ? netProfit / totalIncome : 0;

                // 替换占位符
                text = text.Replace("{TOTAL_INCOME}", $"{totalIncome:F2} 元");
                text = text.Replace("{TOTAL_EXPENSE}", $"{totalExpense:F2} 元");
                text = text.Replace("{NET_PROFIT}", $"{netProfit:F2} 元");
                text = text.Replace("{PROFIT_MARGIN}", $"{profitMargin:P2}");

                range.Text = text;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充财务总结信息时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 生成示例项目进度数据
        /// </summary>
        /// <returns>项目进度数据列表</returns>
        private List<ProjectProgressData> GenerateSampleProjectProgressData()
        {
            var startDate = DateTime.Now.AddMonths(-2);
            return new List<ProjectProgressData>
            {
                new ProjectProgressData {
                    TaskName = "需求分析", Assignee = "张三",
                    PlannedStart = startDate, ActualStart = startDate,
                    PlannedEnd = startDate.AddDays(10), EstimatedEnd = startDate.AddDays(10),
                    CompletionPercentage = 100, Status = "已完成"
                },
                new ProjectProgressData {
                    TaskName = "系统设计", Assignee = "李四",
                    PlannedStart = startDate.AddDays(10), ActualStart = startDate.AddDays(10),
                    PlannedEnd = startDate.AddDays(20), EstimatedEnd = startDate.AddDays(20),
                    CompletionPercentage = 100, Status = "已完成"
                },
                new ProjectProgressData {
                    TaskName = "前端开发", Assignee = "王五",
                    PlannedStart = startDate.AddDays(20), ActualStart = startDate.AddDays(20),
                    PlannedEnd = startDate.AddDays(40), EstimatedEnd = startDate.AddDays(42),
                    CompletionPercentage = 80, Status = "进行中"
                },
                new ProjectProgressData {
                    TaskName = "后端开发", Assignee = "赵六",
                    PlannedStart = startDate.AddDays(20), ActualStart = startDate.AddDays(20),
                    PlannedEnd = startDate.AddDays(40), EstimatedEnd = startDate.AddDays(40),
                    CompletionPercentage = 90, Status = "进行中"
                },
                new ProjectProgressData {
                    TaskName = "系统测试", Assignee = "孙七",
                    PlannedStart = startDate.AddDays(40), ActualStart = startDate.AddDays(42),
                    PlannedEnd = startDate.AddDays(50), EstimatedEnd = startDate.AddDays(52),
                    CompletionPercentage = 30, Status = "进行中"
                }
            };
        }

        /// <summary>
        /// 填充项目进度数据
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="projectData">项目进度数据</param>
        private void FillProjectProgressData(IWordDocument document, List<ProjectProgressData> projectData)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                int tablePosition = text.IndexOf("{PROGRESS_DETAILS}");
                if (tablePosition >= 0)
                {
                    var tableRange = document.Range(tablePosition, tablePosition + 18);
                    var table = document.Tables.Add(tableRange, projectData.Count + 1, 7);

                    // 设置表头
                    table.Cell(1, 1).Range.Text = "任务名称";
                    table.Cell(1, 2).Range.Text = "负责人";
                    table.Cell(1, 3).Range.Text = "计划开始";
                    table.Cell(1, 4).Range.Text = "实际开始";
                    table.Cell(1, 5).Range.Text = "完成度";
                    table.Cell(1, 6).Range.Text = "预计完成";
                    table.Cell(1, 7).Range.Text = "状态";

                    // 格式化表头
                    for (int i = 1; i <= 7; i++)
                    {
                        var cell = table.Cell(1, i);
                        cell.Range.Font.Bold = 1;
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                    }

                    // 填充数据
                    for (int i = 0; i < projectData.Count; i++)
                    {
                        var data = projectData[i];
                        table.Cell(i + 2, 1).Range.Text = data.TaskName;
                        table.Cell(i + 2, 2).Range.Text = data.Assignee;
                        table.Cell(i + 2, 3).Range.Text = data.PlannedStart.ToString("MM-dd");
                        table.Cell(i + 2, 4).Range.Text = data.ActualStart.ToString("MM-dd");
                        table.Cell(i + 2, 5).Range.Text = $"{data.CompletionPercentage}%";
                        table.Cell(i + 2, 6).Range.Text = data.EstimatedEnd.ToString("MM-dd");
                        table.Cell(i + 2, 7).Range.Text = data.Status;

                        // 格式化数据行
                        for (int j = 1; j <= 7; j++)
                        {
                            table.Cell(i + 2, j).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        }

                        // 根据状态设置颜色
                        switch (data.Status)
                        {
                            case "已完成":
                                table.Cell(i + 2, 7).Shading.BackgroundPatternColor = WdColor.wdColorBrightGreen;
                                break;
                            case "进行中":
                                table.Cell(i + 2, 7).Shading.BackgroundPatternColor = WdColor.wdColorYellow;
                                break;
                            case "延期":
                                table.Cell(i + 2, 7).Shading.BackgroundPatternColor = WdColor.wdColorRed;
                                break;
                        }
                    }

                    // 设置表格样式
                    table.Borders.Enable = 1;
                    table.AllowAutoFit = true;
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充项目进度数据时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充里程碑
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FillMilestones(IWordDocument document)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                int milestonePosition = text.IndexOf("{MILESTONES}");
                if (milestonePosition >= 0)
                {
                    var milestoneRange = document.Range(milestonePosition, milestonePosition + 12);
                    milestoneRange.Text = "";

                    // 添加里程碑列表
                    var listRange = document.Range(milestonePosition, milestonePosition);
                    listRange.Text = "• 项目启动 - 已完成\n";
                    listRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    listRange.Text = "• 需求分析完成 - 已完成\n";
                    listRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    listRange.Text = "• 系统设计完成 - 已完成\n";
                    listRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    listRange.Text = "• 核心功能开发 - 进行中\n";
                    listRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    listRange.Text = "• 系统测试 - 计划中\n";
                    listRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    listRange.Text = "• 项目交付 - 计划中\n";

                    listRange.ListFormat.ApplyBulletDefault();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充里程碑时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充风险和问题
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void FillRisksAndIssues(IWordDocument document)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                int risksPosition = text.IndexOf("{RISKS_AND_ISSUES}");
                if (risksPosition >= 0)
                {
                    var risksRange = document.Range(risksPosition, risksPosition + 18);
                    risksRange.Text = "";

                    // 添加风险和问题列表
                    var listRange = document.Range(risksPosition, risksPosition);
                    listRange.Text = "• 人员变动风险 - 中等风险，已制定应对措施\n";
                    listRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    listRange.Text = "• 技术难点 - 低风险，正在解决中\n";
                    listRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    listRange.Text = "• 进度延期 - 低风险，已调整计划\n";

                    listRange.ListFormat.ApplyBulletDefault();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充风险和问题时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充项目总结信息
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="projectData">项目进度数据</param>
        private void FillProjectSummaryInfo(IWordDocument document, List<ProjectProgressData> projectData)
        {
            try
            {
                var range = document.Range();
                var text = range.Text;

                // 计算总结数据
                int completedTasks = projectData.Count(d => d.Status == "已完成");
                int totalTasks = projectData.Count;
                int progressPercentage = totalTasks > 0 ? (completedTasks * 100 / totalTasks) : 0;
                string projectStatus = progressPercentage >= 80 ? "绿色" : progressPercentage >= 50 ? "黄色" : "红色";
                DateTime estimatedCompletion = projectData.Max(d => d.EstimatedEnd);
                decimal budgetUsage = 75.5m; // 示例数据

                // 替换占位符
                text = text.Replace("{PROJECT_STATUS}", projectStatus);
                text = text.Replace("{ESTIMATED_COMPLETION}", estimatedCompletion.ToString("yyyy年MM月dd日"));
                text = text.Replace("{BUDGET_USAGE}", $"{budgetUsage:F1}%");

                range.Text = text;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"填充项目总结信息时出错: {ex.Message}");
            }
        }
    }
}