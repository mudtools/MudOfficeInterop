//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace EnterpriseReportingSystemSample
{
    /// <summary>
    /// 企业报表生成系统示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel构建企业级报表系统
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("企业报表生成系统示例");
            Console.WriteLine("===============");
            Console.WriteLine();

            // 演示报表模板设计
            ReportTemplateDesignExample();

            // 演示数据填充与合并
            DataPopulationAndMergeExample();

            // 演示报表格式设置
            ReportFormattingExample();

            // 演示批量报表生成
            BatchReportGenerationExample();

            // 演示报表导出功能
            ReportExportExample();

            // 演示报表验证系统
            ReportValidationExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 报表模板设计示例
        /// 演示如何设计和使用报表模板
        /// </summary>
        static void ReportTemplateDesignExample()
        {
            Console.WriteLine("=== 报表模板设计示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "销售报表模板";

                // 设计报表模板结构
                // 标题行
                worksheet.Range("A1").Value = "XYZ公司月度销售报表";
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.Navy;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1:F1").Merge();

                // 报表信息行
                worksheet.Range("A2").Value = "报表期间:";
                worksheet.Range("A2").Font.Bold = true;
                worksheet.Range("B2").Value = "{报表期间}"; // 占位符

                worksheet.Range("D2").Value = "生成时间:";
                worksheet.Range("D2").Font.Bold = true;
                worksheet.Range("E2").Value = "{生成时间}"; // 占位符

                // 表头
                worksheet.Range("A4").Value = "产品ID";
                worksheet.Range("B4").Value = "产品名称";
                worksheet.Range("C4").Value = "销售数量";
                worksheet.Range("D4").Value = "单价";
                worksheet.Range("E4").Value = "总金额";
                worksheet.Range("F4").Value = "销售占比";

                // 设置表头格式
                var headerRange = worksheet.Range("A4:F4");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 数据行占位符
                worksheet.Range("A5").Value = "{产品ID}";
                worksheet.Range("B5").Value = "{产品名称}";
                worksheet.Range("C5").Value = "{销售数量}";
                worksheet.Range("D5").Value = "{单价}";
                worksheet.Range("E5").Value = "{总金额}";
                worksheet.Range("F5").Value = "{销售占比}";

                // 设置数据行格式
                var dataRange = worksheet.Range("A5:F5");
                dataRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                dataRange.NumberFormat = "0"; // 数字格式占位符

                // 总计行
                worksheet.Range("A10").Value = "总计";
                worksheet.Range("A10").Font.Bold = true;
                worksheet.Range("C10").Formula = "=SUM(C5:C9)";
                worksheet.Range("E10").Formula = "=SUM(E5:E9)";

                // 设置总计行格式
                var totalRange = worksheet.Range("A10:F10");
                totalRange.Font.Bold = true;
                totalRange.Interior.Color = Color.LightBlue;
                totalRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 图表区域
                worksheet.Range("H1").Value = "销售图表";
                worksheet.Range("H1").Font.Bold = true;
                worksheet.Range("H1:L15").Interior.Color = Color.LightYellow;
                worksheet.Range("H1:L15").Borders.LineStyle = XlLineStyle.xlContinuous;

                // 页脚
                worksheet.Range("A12").Value = "制表人: {制表人}";
                worksheet.Range("D12").Value = "审核人: {审核人}";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存模板文件
                string templateFileName = $"SalesReportTemplate_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(templateFileName);

                Console.WriteLine($"✓ 成功设计报表模板: {templateFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 报表模板设计时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据填充与合并示例
        /// 演示如何将数据填充到报表模板中
        /// </summary>
        static void DataPopulationAndMergeExample()
        {
            Console.WriteLine("=== 数据填充与合并示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "销售数据报表";

                // 创建报表标题
                worksheet.Range("A1").Value = "ABC公司2023年第三季度销售报表";
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.DarkGreen;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1:G1").Merge();

                // 报表信息
                worksheet.Range("A2").Value = "报表期间:";
                worksheet.Range("A2").Font.Bold = true;
                worksheet.Range("B2").Value = "2023年7月-9月";

                worksheet.Range("E2").Value = "生成时间:";
                worksheet.Range("E2").Font.Bold = true;
                worksheet.Range("F2").Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // 表头
                worksheet.Range("A4").Value = "产品ID";
                worksheet.Range("B4").Value = "产品名称";
                worksheet.Range("C4").Value = "7月销量";
                worksheet.Range("D4").Value = "8月销量";
                worksheet.Range("E4").Value = "9月销量";
                worksheet.Range("F4").Value = "季度总销量";
                worksheet.Range("G4").Value = "平均月销量";

                // 设置表头格式
                var headerRange = worksheet.Range("A4:G4");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 销售数据
                object[,] salesData = {
                    {"P001", "笔记本电脑", 120, 150, 180, "=SUM(C5:E5)", "=AVERAGE(C5:E5)"},
                    {"P002", "台式电脑", 80, 90, 100, "=SUM(C6:E6)", "=AVERAGE(C6:E6)"},
                    {"P003", "平板电脑", 200, 220, 250, "=SUM(C7:E7)", "=AVERAGE(C7:E7)"},
                    {"P004", "智能手机", 500, 550, 600, "=SUM(C8:E8)", "=AVERAGE(C8:E8)"},
                    {"P005", "智能手表", 300, 320, 350, "=SUM(C9:E9)", "=AVERAGE(C9:E9)"}
                };

                var dataRange = worksheet.Range("A5:G9");
                dataRange.Value = salesData;

                // 设置数据格式
                worksheet.Range("A5:B9").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range("C5:G9").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range("C5:G9").NumberFormat = "0";

                // 总计行
                worksheet.Range("A11").Value = "总计";
                worksheet.Range("A11").Font.Bold = true;
                worksheet.Range("C11").Formula = "=SUM(C5:C9)";
                worksheet.Range("D11").Formula = "=SUM(D5:D9)";
                worksheet.Range("E11").Formula = "=SUM(E5:E9)";
                worksheet.Range("F11").Formula = "=SUM(F5:F9)";
                worksheet.Range("G11").Formula = "=AVERAGE(G5:G9)";

                // 设置总计行格式
                var totalRange = worksheet.Range("A11:G11");
                totalRange.Font.Bold = true;
                totalRange.Interior.Color = Color.LightBlue;
                totalRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 添加条件格式
                var conditionalRange = worksheet.Range("F5:F9");
                var condition = conditionalRange.FormatConditions.Add(
                    XlFormatConditionType.xlCellValue,
                    XlFormatConditionOperator.xlGreater,
                    1000);
                condition.Interior.Color = Color.LightGreen;

                // 页脚信息
                worksheet.Range("A13").Value = "制表人: 张三";
                worksheet.Range("D13").Value = "审核人: 李四";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存报表文件
                string reportFileName = $"SalesDataReport_{DateTime.Now:yyyyMMddHHmmss}.xlsx";

                workbook.SaveAs(reportFileName);

                Console.WriteLine($"✓ 成功演示数据填充与合并: {reportFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数据填充与合并时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 报表格式设置示例
        /// 演示如何设置报表格式
        /// </summary>
        static void ReportFormattingExample()
        {
            Console.WriteLine("=== 报表格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "格式化报表";

                // 创建报表标题
                worksheet.Range("A1").Value = "DEF公司财务分析报表";
                worksheet.Range("A1").Font.Size = 18;
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1").Interior.Color = Color.DarkBlue;
                worksheet.Range("A1:J1").Merge();
                worksheet.Range("A1:J1").HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 报表信息
                worksheet.Range("A2").Value = "报表期间:";
                worksheet.Range("B2").Value = "2023年1月-9月";
                worksheet.Range("B2").Font.Bold = true;

                worksheet.Range("H2").Value = "单位: 万元";
                worksheet.Range("H2").Font.Italic = true;

                // 创建多级表头
                worksheet.Range("A4").Value = "财务指标";
                worksheet.Range("A4").Font.Bold = true;
                worksheet.Range("A4").Interior.Color = Color.LightGray;
                worksheet.Range("A4").HorizontalAlignment = XlHAlign.xlHAlignCenter;
                worksheet.Range("A4:A5").Merge();

                worksheet.Range("B4").Value = "第一季度";
                worksheet.Range("B4:D4").Merge();
                worksheet.Range("B4").Interior.Color = Color.LightBlue;
                worksheet.Range("B4").HorizontalAlignment = XlHAlign.xlHAlignCenter;

                worksheet.Range("E4").Value = "第二季度";
                worksheet.Range("E4:G4").Merge();
                worksheet.Range("E4").Interior.Color = Color.LightGreen;
                worksheet.Range("E4").HorizontalAlignment = XlHAlign.xlHAlignCenter;

                worksheet.Range("H4").Value = "第三季度";
                worksheet.Range("H4:J4").Merge();
                worksheet.Range("H4").Interior.Color = Color.LightYellow;
                worksheet.Range("H4").HorizontalAlignment = XlHAlign.xlHAlignCenter;

                worksheet.Range("B5").Value = "1月";
                worksheet.Range("C5").Value = "2月";
                worksheet.Range("D5").Value = "3月";
                worksheet.Range("E5").Value = "4月";
                worksheet.Range("F5").Value = "5月";
                worksheet.Range("G5").Value = "6月";
                worksheet.Range("H5").Value = "7月";
                worksheet.Range("I5").Value = "8月";
                worksheet.Range("J5").Value = "9月";

                // 设置二级表头格式
                var subHeaderRange = worksheet.Range("B5:J5");
                subHeaderRange.Font.Bold = true;
                subHeaderRange.Interior.Color = Color.LightGray;
                subHeaderRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 财务数据
                object[,] financialData = {
                    {"营业收入", 1200, 1300, 1400, 1500, 1600, 1700, 1800, 1900, 2000},
                    {"营业成本", 800, 850, 900, 950, 1000, 1050, 1100, 1150, 1200},
                    {"毛利润", 400, 450, 500, 550, 600, 650, 700, 750, 800},
                    {"毛利率(%)", 33.33, 34.62, 35.71, 36.67, 37.50, 38.24, 38.89, 39.47, 40.00}
                };

                var dataRange = worksheet.Range("A6:J9");
                dataRange.Value = financialData;

                // 设置数据格式
                worksheet.Range("A6:A9").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range("B6:J9").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range("B6:J8").NumberFormat = "0";
                worksheet.Range("B9:J9").NumberFormat = "0.00";

                // 添加数据条条件格式
                var dataBarRange = worksheet.Range("B6:J8");
                var dataBar = dataBarRange.FormatConditions.AddDatabar();
                dataBar.BarColor.Color = Color.Blue;

                // 总计行
                worksheet.Range("A11").Value = "总计";
                worksheet.Range("A11").Font.Bold = true;

                for (int col = 2; col <= 10; col++)
                {
                    string colLetter = GetColumnLetter(col);
                    worksheet.Range($"{colLetter}11").Formula = $"=SUM({colLetter}6:{colLetter}9)";
                }

                // 设置总计行格式
                var totalRange = worksheet.Range("A11:J11");
                totalRange.Font.Bold = true;
                totalRange.Interior.Color = Color.Gold;
                totalRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 添加图表
                var chartRange = worksheet.Range("A6:J9");
                var chart = worksheet.Shapes.AddChart2();
                chart.Chart.SetSourceData(chartRange);
                chart.Chart.ChartType = MsoChartType.xlLineMarkers;
                chart.Name = "财务趋势图";
                chart.Top = Convert.ToSingle(worksheet.Range("A13").Top);
                chart.Left = Convert.ToSingle(worksheet.Range("A13").Left);
                chart.Width = 800;
                chart.Height = 300;

                // 页脚信息
                worksheet.Range("A18").Value = "制表人: 王五";
                worksheet.Range("F18").Value = "审核人: 赵六";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存报表文件
                string formattedReportFileName = $"FormattedReport_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(formattedReportFileName);

                Console.WriteLine($"✓ 成功演示报表格式设置: {formattedReportFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 报表格式设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 批量报表生成示例
        /// 演示如何批量生成报表
        /// </summary>
        static void BatchReportGenerationExample()
        {
            Console.WriteLine("=== 批量报表生成示例 ===");

            try
            {
                // 模拟部门数据
                var departments = new List<string> { "销售部", "市场部", "技术部", "人事部", "财务部" };

                foreach (var department in departments)
                {
                    // 为每个部门创建报表
                    using var excelApp = ExcelFactory.BlankWorkbook();

                    // 获取活动工作簿和工作表
                    var workbook = excelApp.ActiveWorkbook;
                    var worksheet = workbook.ActiveSheetWrap;
                    worksheet.Name = $"{department}报表";

                    // 创建报表标题
                    worksheet.Range("A1").Value = $"{department}月度绩效报表";
                    worksheet.Range("A1").Font.Size = 16;
                    worksheet.Range("A1").Font.Bold = true;
                    worksheet.Range("A1").Interior.Color = Color.DarkRed;
                    worksheet.Range("A1").Font.Color = Color.White;
                    worksheet.Range("A1:G1").Merge();
                    worksheet.Range("A1:G1").HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    // 报表信息
                    worksheet.Range("A2").Value = "部门:";
                    worksheet.Range("B2").Value = department;
                    worksheet.Range("B2").Font.Bold = true;

                    worksheet.Range("E2").Value = "报表月份:";
                    worksheet.Range("F2").Value = "2023年9月";
                    worksheet.Range("F2").Font.Bold = true;

                    // 表头
                    worksheet.Range("A4").Value = "员工姓名";
                    worksheet.Range("B4").Value = "岗位";
                    worksheet.Range("C4").Value = "基本工资";
                    worksheet.Range("D4").Value = "绩效奖金";
                    worksheet.Range("E4").Value = "加班费";
                    worksheet.Range("F4").Value = "应发工资";
                    worksheet.Range("G4").Value = "绩效评分";

                    // 设置表头格式
                    var headerRange = worksheet.Range("A4:G4");
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = Color.LightGray;
                    headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    // 生成员工数据
                    var employees = GenerateEmployeeData(department);
                    var dataRange = worksheet.Range($"A5:G{4 + employees.GetLength(0)}");
                    dataRange.Value = employees;

                    // 设置数据格式
                    worksheet.Range("A5:B" + (4 + employees.GetLength(0))).Borders.LineStyle = XlLineStyle.xlContinuous;
                    worksheet.Range("C5:G" + (4 + employees.GetLength(0))).Borders.LineStyle = XlLineStyle.xlContinuous;
                    worksheet.Range("C5:F" + (4 + employees.GetLength(0))).NumberFormat = "¥#,##0";
                    worksheet.Range("G5:G" + (4 + employees.GetLength(0))).NumberFormat = "0.0";

                    // 总计行
                    int lastDataRow = 4 + employees.GetLength(0);
                    worksheet.Range("A" + (lastDataRow + 1)).Value = "部门总计";
                    worksheet.Range("A" + (lastDataRow + 1)).Font.Bold = true;

                    worksheet.Range("C" + (lastDataRow + 1)).Formula = $"=SUM(C5:C{lastDataRow})";
                    worksheet.Range("D" + (lastDataRow + 1)).Formula = $"=SUM(D5:D{lastDataRow})";
                    worksheet.Range("E" + (lastDataRow + 1)).Formula = $"=SUM(E5:E{lastDataRow})";
                    worksheet.Range("F" + (lastDataRow + 1)).Formula = $"=SUM(F5:F{lastDataRow})";
                    worksheet.Range("G" + (lastDataRow + 1)).Formula = $"=AVERAGE(G5:G{lastDataRow})";

                    // 设置总计行格式
                    var totalRange = worksheet.Range("A" + (lastDataRow + 1) + ":G" + (lastDataRow + 1));
                    totalRange.Font.Bold = true;
                    totalRange.Interior.Color = Color.LightBlue;
                    totalRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                    // 添加条件格式突出显示高绩效员工
                    var performanceRange = worksheet.Range("G5:G" + lastDataRow);
                    var condition = performanceRange?.FormatConditions.Add(
                        XlFormatConditionType.xlCellValue,
                        XlFormatConditionOperator.xlGreater,
                        4.5);
                    condition.Interior.Color = Color.LightGreen;

                    // 页脚信息
                    worksheet.Range("A" + (lastDataRow + 3)).Value = "制表人: HR部门";
                    worksheet.Range("D" + (lastDataRow + 3)).Value = "审核人: 总经理";

                    // 自动调整列宽
                    worksheet.Columns.AutoFit();

                    // 保存报表文件
                    string batchReportFileName = $"{department}Report_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    workbook.SaveAs(batchReportFileName);

                    Console.WriteLine($"  ✓ 已生成 {department} 报表: {batchReportFileName}");
                }

                Console.WriteLine($"✓ 成功演示批量报表生成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 批量报表生成时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 报表导出示例
        /// 演示如何导出报表为不同格式
        /// </summary>
        static void ReportExportExample()
        {
            Console.WriteLine("=== 报表导出示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "导出报表";

                // 创建报表内容
                worksheet.Range("A1").Value = "GHI公司年度总结报表";
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.Purple;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1:H1").Merge();
                worksheet.Range("A1:H1").HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 报表信息
                worksheet.Range("A2").Value = "报表年度:";
                worksheet.Range("B2").Value = "2023年";
                worksheet.Range("B2").Font.Bold = true;

                worksheet.Range("F2").Value = "生成日期:";
                worksheet.Range("G2").Value = DateTime.Now.ToString("yyyy-MM-dd");
                worksheet.Range("G2").Font.Bold = true;

                // 表头
                worksheet.Range("A4").Value = "业务板块";
                worksheet.Range("B4").Value = "Q1收入";
                worksheet.Range("C4").Value = "Q2收入";
                worksheet.Range("D4").Value = "Q3收入";
                worksheet.Range("E4").Value = "Q4收入";
                worksheet.Range("F4").Value = "年度总收入";
                worksheet.Range("G4").Value = "同比增长";
                worksheet.Range("H4").Value = "市场份额";

                // 设置表头格式
                var headerRange = worksheet.Range("A4:H4");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 业务数据
                object[,] businessData = {
                    {"电子产品", 12000, 13500, 14800, 16200, "=SUM(B5:E5)", "12.5%", "25.3%"},
                    {"家居用品", 8000, 8500, 9200, 9800, "=SUM(B6:E6)", "8.7%", "18.6%"},
                    {"服装服饰", 6500, 7200, 7800, 8400, "=SUM(B7:E7)", "10.2%", "15.2%"},
                    {"食品饮料", 9200, 9800, 10500, 11200, "=SUM(B8:E8)", "9.8%", "20.1%"},
                    {"运动健康", 5400, 6100, 6700, 7300, "=SUM(B9:E9)", "12.3%", "10.8%"}
                };

                var dataRange = worksheet.Range("A5:H9");
                dataRange.Value = businessData;

                // 设置数据格式
                worksheet.Range("A5:A9").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range("B5:H9").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range("B5:F9").NumberFormat = "¥#,##0";
                worksheet.Range("G5:G9").NumberFormat = "0.0%";
                worksheet.Range("H5:H9").NumberFormat = "0.0%";

                // 总计行
                worksheet.Range("A11").Value = "公司总计";
                worksheet.Range("A11").Font.Bold = true;

                for (int col = 2; col <= 6; col++)
                {
                    string colLetter = GetColumnLetter(col);
                    worksheet.Range($"{colLetter}11").Formula = $"=SUM({colLetter}5:{colLetter}9)";
                }

                worksheet.Range("G11").Formula = "=AVERAGE(G5:G9)";
                worksheet.Range("H11").Formula = "=SUM(H5:H9)";

                // 设置总计行格式
                var totalRange = worksheet.Range("A11:H11");
                totalRange.Font.Bold = true;
                totalRange.Interior.Color = Color.Gold;
                totalRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 添加图表
                var chartRange = worksheet.Range("A5:F9");
                var chart = worksheet.Shapes.AddChart2();
                chart.Chart.SetSourceData(chartRange);
                chart.Chart.ChartType = MsoChartType.xlColumnClustered;
                chart.Name = "各业务板块收入对比";
                chart.Top = Convert.ToSingle(worksheet.Range("A13").Top);
                chart.Left = Convert.ToSingle(worksheet.Range("A13").Left);
                chart.Width = 700;
                chart.Height = 300;

                // 页脚信息
                worksheet.Range("A18").Value = "制表人: 财务部";
                worksheet.Range("E18").Value = "审核人: 董事会";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存为Excel格式
                string excelFileName = $"ExportReport_Excel_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(excelFileName);

                // 注意：实际导出为PDF、CSV等格式需要特定的Excel功能支持
                // 这里仅演示概念，实际导出需要Excel应用程序支持

                Console.WriteLine($"✓ 成功演示报表导出功能");
                Console.WriteLine($"  Excel格式: {excelFileName}");
                Console.WriteLine("  注意：PDF、CSV等格式导出需要Excel应用程序支持");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 报表导出时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 报表验证示例
        /// 演示如何验证报表数据和格式
        /// </summary>
        static void ReportValidationExample()
        {
            Console.WriteLine("=== 报表验证示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "验证报表";

                // 创建报表标题
                worksheet.Range("A1").Value = "JKL公司数据验证报表";
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.Orange;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1:G1").Merge();
                worksheet.Range("A1:G1").HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 报表信息
                worksheet.Range("A2").Value = "验证日期:";
                worksheet.Range("B2").Value = DateTime.Now.ToString("yyyy-MM-dd");
                worksheet.Range("B2").Font.Bold = true;

                worksheet.Range("E2").Value = "验证状态:";
                worksheet.Range("F2").Value = "待验证";
                worksheet.Range("F2").Font.Bold = true;

                // 表头
                worksheet.Range("A4").Value = "数据项";
                worksheet.Range("B4").Value = "预期值范围";
                worksheet.Range("C4").Value = "实际值";
                worksheet.Range("D4").Value = "验证结果";
                worksheet.Range("E4").Value = "偏差率";
                worksheet.Range("F4").Value = "备注";

                // 设置表头格式
                var headerRange = worksheet.Range("A4:F4");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 验证数据
                object[,] validationData = {
                    {"销售额", "10000-50000", 35000, "通过", "0%", ""},
                    {"利润率", "10%-25%", 0.18, "通过", "0%", ""},
                    {"客户满意度", "85%-100%", 0.92, "通过", "0%", ""},
                    {"市场份额", "15%-30%", 0.22, "通过", "0%", ""},
                    {"员工满意度", "75%-95%", 0.88, "通过", "0%", ""},
                    {"研发投入占比", "5%-15%", 0.12, "通过", "0%", ""}
                };

                var dataRange = worksheet.Range("A5:F10");
                dataRange.Value = validationData;

                // 设置数据格式
                worksheet.Range("A5:F10").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range("C5:C10").NumberFormat = "0";
                worksheet.Range("D5:D10").Interior.Color = Color.LightGreen;
                worksheet.Range("E5:E10").NumberFormat = "0.00%";

                // 添加数据验证规则
                var validation = worksheet.Range("C5").Validation;
                validation.Add(XlDVType.xlValidateWholeNumber, XlDVAlertStyle.xlValidAlertStop,
                    XlFormatConditionOperator.xlBetween, 10000, 50000);
                validation.InputTitle = "销售额输入";
                validation.InputMessage = "请输入10000-50000之间的数值";
                validation.ErrorTitle = "输入错误";
                validation.ErrorMessage = "销售额必须在10000-50000之间";

                // 验证结果统计
                worksheet.Range("A12").Value = "验证统计";
                worksheet.Range("A12").Font.Bold = true;

                worksheet.Range("A13").Value = "总项数:";
                worksheet.Range("B13").Value = 6;

                worksheet.Range("A14").Value = "通过项数:";
                worksheet.Range("B14").Value = 6;

                worksheet.Range("A15").Value = "通过率:";
                worksheet.Range("B15").Formula = "=B14/B13";
                worksheet.Range("B15").NumberFormat = "0.00%";
                worksheet.Range("B15").Interior.Color = Color.LightBlue;

                // 设置统计区域格式
                var statsRange = worksheet.Range("A12:B15");
                statsRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 验证状态更新
                worksheet.Range("F2").Value = "验证完成";
                worksheet.Range("F2").Interior.Color = Color.LightGreen;

                // 页脚信息
                worksheet.Range("A17").Value = "验证人: 数据分析师";
                worksheet.Range("D17").Value = "复核人: 质量管理部门";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存报表文件
                string validationReportFileName = $"ValidationReport_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(validationReportFileName);

                Console.WriteLine($"✓ 成功演示报表验证功能: {validationReportFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 报表验证时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 生成员工数据
        /// </summary>
        /// <param name="department">部门名称</param>
        /// <returns>员工数据数组</returns>
        static object[,] GenerateEmployeeData(string department)
        {
            var random = new Random();
            var employees = new List<(string name, string position, int baseSalary, int bonus, int overtime)>();

            switch (department)
            {
                case "销售部":
                    employees.AddRange(new[] {
                        ("张三", "销售经理", 12000, 5000, 800),
                        ("李四", "高级销售", 8000, 3000, 500),
                        ("王五", "销售专员", 6000, 2000, 300),
                        ("赵六", "销售助理", 5000, 1000, 200)
                    });
                    break;

                case "市场部":
                    employees.AddRange(new[] {
                        ("钱七", "市场总监", 15000, 6000, 1000),
                        ("孙八", "市场经理", 10000, 4000, 600),
                        ("周九", "市场专员", 7000, 2500, 400),
                        ("吴十", "推广专员", 6000, 2000, 300)
                    });
                    break;

                case "技术部":
                    employees.AddRange(new[] {
                        ("郑一", "技术总监", 20000, 8000, 1500),
                        ("王二", "高级工程师", 15000, 6000, 1000),
                        ("冯三", "软件工程师", 12000, 4000, 800),
                        ("陈四", "测试工程师", 10000, 3000, 500),
                        ("褚五", "运维工程师", 11000, 3500, 600)
                    });
                    break;

                case "人事部":
                    employees.AddRange(new[] {
                        ("卫六", "人事总监", 14000, 5000, 800),
                        ("蒋七", "招聘经理", 10000, 3500, 500),
                        ("沈八", "培训专员", 8000, 2500, 300)
                    });
                    break;

                case "财务部":
                    employees.AddRange(new[] {
                        ("韩九", "财务总监", 16000, 6000, 1000),
                        ("杨十", "财务经理", 12000, 4500, 600),
                        ("朱一", "会计", 8000, 2000, 300),
                        ("秦二", "出纳", 6000, 1500, 200)
                    });
                    break;
            }

            var data = new object[employees.Count, 7];
            for (int i = 0; i < employees.Count; i++)
            {
                data[i, 0] = employees[i].name;
                data[i, 1] = employees[i].position;
                data[i, 2] = employees[i].baseSalary;
                data[i, 3] = employees[i].bonus;
                data[i, 4] = employees[i].overtime;
                data[i, 5] = $"=C{i + 5}+D{i + 5}+E{i + 5}"; // 应发工资公式
                data[i, 6] = Math.Round(3.0 + random.NextDouble() * 2.0, 1); // 绩效评分 3.0-5.0
            }

            return data;
        }

        /// <summary>
        /// 根据列号获取列字母
        /// </summary>
        /// <param name="columnNumber">列号</param>
        /// <returns>列字母</returns>
        static string GetColumnLetter(int columnNumber)
        {
            string columnLetter = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnLetter = (char)(65 + modulo) + columnLetter;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnLetter;
        }
    }
}