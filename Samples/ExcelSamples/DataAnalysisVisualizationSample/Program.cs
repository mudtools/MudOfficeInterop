//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace DataAnalysisVisualizationSample
{
    /// <summary>
    /// 数据分析与可视化示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行数据分析和可视化操作
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("数据分析与可视化示例");
            Console.WriteLine("====================");
            Console.WriteLine();

            // 演示描述性统计分析
            DescriptiveStatisticsExample();

            // 演示趋势分析
            TrendAnalysisExample();

            // 演示相关性分析
            CorrelationAnalysisExample();

            // 演示频率分布分析
            FrequencyDistributionExample();

            // 演示数据可视化
            DataVisualizationExample();

            // 演示综合数据分析报告
            ComprehensiveAnalysisReportExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 描述性统计分析示例
        /// 演示如何计算和展示描述性统计数据
        /// </summary>
        static void DescriptiveStatisticsExample()
        {
            Console.WriteLine("=== 描述性统计分析示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "描述性统计";

                // 创建销售数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "成本";
                worksheet.Range("D1").Value = "利润";

                object[,] salesData = {
                    {"1月", 100000, 70000, 30000},
                    {"2月", 120000, 80000, 40000},
                    {"3月", 140000, 90000, 50000},
                    {"4月", 130000, 85000, 45000},
                    {"5月", 150000, 95000, 55000},
                    {"6月", 160000, 100000, 60000},
                    {"7月", 170000, 105000, 65000},
                    {"8月", 180000, 110000, 70000},
                    {"9月", 190000, 115000, 75000},
                    {"10月", 200000, 120000, 80000},
                    {"11月", 210000, 125000, 85000},
                    {"12月", 220000, 130000, 90000}
                };

                var dataRange = worksheet.Range("A2:D13");
                dataRange.Value = salesData;

                // 计算描述性统计
                // 销售额统计
                worksheet.Range("F1").Value = "销售额统计";
                worksheet.Range("F1").Font.Bold = true;
                worksheet.Range("F1").Interior.Color = Color.LightBlue;

                worksheet.Range("F2").Value = "总和";
                worksheet.Range("G2").Formula = "=SUM(B2:B13)";

                worksheet.Range("F3").Value = "平均值";
                worksheet.Range("G3").Formula = "=AVERAGE(B2:B13)";

                worksheet.Range("F4").Value = "中位数";
                worksheet.Range("G4").Formula = "=MEDIAN(B2:B13)";

                worksheet.Range("F5").Value = "最大值";
                worksheet.Range("G5").Formula = "=MAX(B2:B13)";

                worksheet.Range("F6").Value = "最小值";
                worksheet.Range("G6").Formula = "=MIN(B2:B13)";

                worksheet.Range("F7").Value = "标准差";
                worksheet.Range("G7").Formula = "=STDEV(B2:B13)";

                worksheet.Range("F8").Value = "方差";
                worksheet.Range("G8").Formula = "=VAR(B2:B13)";

                // 成本统计
                worksheet.Range("I1").Value = "成本统计";
                worksheet.Range("I1").Font.Bold = true;
                worksheet.Range("I1").Interior.Color = Color.LightGreen;

                worksheet.Range("I2").Value = "总和";
                worksheet.Range("J2").Formula = "=SUM(C2:C13)";

                worksheet.Range("I3").Value = "平均值";
                worksheet.Range("J3").Formula = "=AVERAGE(C2:C13)";

                worksheet.Range("I4").Value = "中位数";
                worksheet.Range("J4").Formula = "=MEDIAN(C2:C13)";

                worksheet.Range("I5").Value = "最大值";
                worksheet.Range("J5").Formula = "=MAX(C2:C13)";

                worksheet.Range("I6").Value = "最小值";
                worksheet.Range("J6").Formula = "=MIN(C2:C13)";

                worksheet.Range("I7").Value = "标准差";
                worksheet.Range("J7").Formula = "=STDEV(C2:C13)";

                worksheet.Range("I8").Value = "方差";
                worksheet.Range("J8").Formula = "=VAR(C2:C13)";

                // 利润统计
                worksheet.Range("L1").Value = "利润统计";
                worksheet.Range("L1").Font.Bold = true;
                worksheet.Range("L1").Interior.Color = Color.LightYellow;

                worksheet.Range("L2").Value = "总和";
                worksheet.Range("M2").Formula = "=SUM(D2:D13)";

                worksheet.Range("L3").Value = "平均值";
                worksheet.Range("M3").Formula = "=AVERAGE(D2:D13)";

                worksheet.Range("L4").Value = "中位数";
                worksheet.Range("M4").Formula = "=MEDIAN(D2:D13)";

                worksheet.Range("L5").Value = "最大值";
                worksheet.Range("M5").Formula = "=MAX(D2:D13)";

                worksheet.Range("L6").Value = "最小值";
                worksheet.Range("M6").Formula = "=MIN(D2:D13)";

                worksheet.Range("L7").Value = "标准差";
                worksheet.Range("M7").Formula = "=STDEV(D2:D13)";

                worksheet.Range("L8").Value = "方差";
                worksheet.Range("M8").Formula = "=VAR(D2:D13)";

                // 设置数字格式
                worksheet.Range("B2:D13").NumberFormat = "¥#,##0";
                worksheet.Range("G2:G8").NumberFormat = "¥#,##0";
                worksheet.Range("J2:J8").NumberFormat = "¥#,##0";
                worksheet.Range("M2:M8").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"DescriptiveStatistics_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示描述性统计分析: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 描述性统计分析时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 趋势分析示例
        /// 演示如何进行数据趋势分析
        /// </summary>
        static void TrendAnalysisExample()
        {
            Console.WriteLine("=== 趋势分析示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "趋势分析";

                // 创建月度销售数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "移动平均";

                object[,] monthlyData = {
                    {"1月", 100000},
                    {"2月", 120000},
                    {"3月", 140000},
                    {"4月", 130000},
                    {"5月", 150000},
                    {"6月", 160000},
                    {"7月", 170000},
                    {"8月", 180000},
                    {"9月", 190000},
                    {"10月", 200000},
                    {"11月", 210000},
                    {"12月", 220000}
                };

                var dataRange = worksheet.Range("A2:B13");
                dataRange.Value = monthlyData;

                // 计算3个月移动平均
                for (int i = 4; i <= 13; i++)
                {
                    worksheet.Range($"C{i}").Formula = $"=AVERAGE(B{i - 2}:B{i})";
                }

                // 添加趋势线分析
                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 600, 400);
                var chart = chartObject.Chart;

                // 设置图表类型为折线图
                chart.ChartType = MsoChartType.xlLine;

                // 设置数据源
                chart.SetSourceData(worksheet.Range("A1:C13"));

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "销售额趋势分析";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 添加趋势线
                var series = chart.SeriesCollection()[1];
                var trendline = series.Trendlines().Add(XlTrendlineType.xlLinear);
                trendline.DisplayEquation = true;
                trendline.DisplayRSquared = true;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "月份";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                // 设置数字格式
                worksheet.Range("B2:B13").NumberFormat = "¥#,##0";
                worksheet.Range("C4:C13").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"TrendAnalysis_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示趋势分析: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 趋势分析时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 相关性分析示例
        /// 演示如何进行变量间的相关性分析
        /// </summary>
        static void CorrelationAnalysisExample()
        {
            Console.WriteLine("=== 相关性分析示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "相关性分析";

                // 创建广告投入与销售额数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "广告投入";
                worksheet.Range("C1").Value = "销售额";
                worksheet.Range("D1").Value = "网站访问量";

                object[,] advertisingData = {
                    {"1月", 5000, 100000, 10000},
                    {"2月", 6000, 120000, 12000},
                    {"3月", 7000, 140000, 14000},
                    {"4月", 6500, 130000, 13000},
                    {"5月", 7500, 150000, 15000},
                    {"6月", 8000, 160000, 16000},
                    {"7月", 8500, 170000, 17000},
                    {"8月", 9000, 180000, 18000},
                    {"9月", 9500, 190000, 19000},
                    {"10月", 10000, 200000, 20000},
                    {"11月", 10500, 210000, 21000},
                    {"12月", 11000, 220000, 22000}
                };

                var dataRange = worksheet.Range("A2:D13");
                dataRange.Value = advertisingData;

                // 计算相关系数
                worksheet.Range("F1").Value = "相关性分析";
                worksheet.Range("F1").Font.Bold = true;
                worksheet.Range("F1").Interior.Color = Color.LightBlue;

                worksheet.Range("F2").Value = "变量";
                worksheet.Range("G2").Value = "广告投入";
                worksheet.Range("H2").Value = "销售额";
                worksheet.Range("I2").Value = "网站访问量";

                worksheet.Range("F3").Value = "广告投入";
                worksheet.Range("F4").Value = "销售额";
                worksheet.Range("F5").Value = "网站访问量";

                // 计算相关系数矩阵
                worksheet.Range("G3").Formula = "=CORREL(B2:B13,B2:B13)"; // 广告投入与广告投入
                worksheet.Range("H3").Formula = "=CORREL(B2:B13,C2:C13)"; // 广告投入与销售额
                worksheet.Range("I3").Formula = "=CORREL(B2:B13,D2:D13)"; // 广告投入与网站访问量

                worksheet.Range("G4").Formula = "=CORREL(C2:C13,B2:B13)"; // 销售额与广告投入
                worksheet.Range("H4").Formula = "=CORREL(C2:C13,C2:C13)"; // 销售额与销售额
                worksheet.Range("I4").Formula = "=CORREL(C2:C13,D2:D13)"; // 销售额与网站访问量

                worksheet.Range("G5").Formula = "=CORREL(D2:D13,B2:B13)"; // 网站访问量与广告投入
                worksheet.Range("H5").Formula = "=CORREL(D2:D13,C2:C13)"; // 网站访问量与销售额
                worksheet.Range("I5").Formula = "=CORREL(D2:D13,D2:D13)"; // 网站访问量与网站访问量

                // 创建散点图分析广告投入与销售额的关系
                var scatterChartObject = worksheet.ChartObjects().Add(300, 50, 500, 350);
                var scatterChart = scatterChartObject.Chart;

                // 设置图表类型为散点图
                scatterChart.ChartType = MsoChartType.xlXYScatter;

                // 设置数据源（广告投入和销售额）
                scatterChart.SetSourceData(worksheet.Range("B1:C13"));

                // 设置标题
                scatterChart.HasTitle = true;
                scatterChart.ChartTitle.Text = "广告投入与销售额关系";

                // 设置图例
                scatterChart.HasLegend = false;

                // 添加趋势线
                var scatterSeries = scatterChart.SeriesCollection()[1];
                var scatterTrendline = scatterSeries.Trendlines().Add(XlTrendlineType.xlLinear);
                scatterTrendline.DisplayEquation = true;
                scatterTrendline.DisplayRSquared = true;

                // 设置坐标轴标题
                scatterChart.Axes(XlAxisType.xlCategory).HasTitle = true;
                scatterChart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "广告投入";

                scatterChart.Axes(XlAxisType.xlValue).HasTitle = true;
                scatterChart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                // 设置数字格式
                worksheet.Range("B2:B13").NumberFormat = "¥#,##0";
                worksheet.Range("C2:C13").NumberFormat = "¥#,##0";
                worksheet.Range("D2:D13").NumberFormat = "0";
                worksheet.Range("G3:I5").NumberFormat = "0.000";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"CorrelationAnalysis_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示相关性分析: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 相关性分析时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 频率分布分析示例
        /// 演示如何进行数据的频率分布分析
        /// </summary>
        static void FrequencyDistributionExample()
        {
            Console.WriteLine("=== 频率分布分析示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "频率分布";

                // 创建销售数据
                worksheet.Range("A1").Value = "销售员";
                worksheet.Range("B1").Value = "销售额";

                object[,] salesData = {
                    {"张三", 120000},
                    {"李四", 150000},
                    {"王五", 90000},
                    {"赵六", 180000},
                    {"钱七", 110000},
                    {"孙八", 200000},
                    {"周九", 80000},
                    {"吴十", 160000},
                    {"郑一", 130000},
                    {"王二", 170000},
                    {"张三", 140000},
                    {"李四", 190000},
                    {"王五", 100000},
                    {"赵六", 210000},
                    {"钱七", 120000},
                    {"孙八", 180000},
                    {"周九", 95000},
                    {"吴十", 165000},
                    {"郑一", 135000},
                    {"王二", 175000}
                };

                var dataRange = worksheet.Range("A2:B21");
                dataRange.Value = salesData;

                // 创建频率分布区间
                worksheet.Range("D1").Value = "区间下限";
                worksheet.Range("E1").Value = "区间上限";
                worksheet.Range("F1").Value = "区间标签";
                worksheet.Range("G1").Value = "频数";
                worksheet.Range("H1").Value = "频率";

                // 定义区间
                int[] lowerBounds = { 80000, 100000, 120000, 140000, 160000, 180000, 200000 };
                int[] upperBounds = { 99999, 119999, 139999, 159999, 179999, 199999, 220000 };
                string[] labels = { "8-10万", "10-12万", "12-14万", "14-16万", "16-18万", "18-20万", "20-22万" };

                for (int i = 0; i < lowerBounds.Length; i++)
                {
                    worksheet.Range($"D{i + 2}").Value = lowerBounds[i];
                    worksheet.Range($"E{i + 2}").Value = upperBounds[i];
                    worksheet.Range($"F{i + 2}").Value = labels[i];
                }

                // 计算频数
                for (int i = 0; i < lowerBounds.Length; i++)
                {
                    worksheet.Range($"G{i + 2}").Formula =
                        $"=COUNTIFS(B2:B21,\">={lowerBounds[i]}\",B2:B21,\"<={upperBounds[i]}\")";
                }

                // 计算频率
                worksheet.Range("H2").Formula = $"=G2/COUNT(B2:B21)";
                for (int i = 1; i < lowerBounds.Length; i++)
                {
                    worksheet.Range($"H{i + 2}").Formula = $"=G{i + 2}/COUNT(B2:B21)";
                }

                // 创建频率分布直方图
                var histogramChartObject = worksheet.ChartObjects().Add(300, 50, 600, 400);
                var histogramChart = histogramChartObject.Chart;

                // 设置图表类型为柱形图
                histogramChart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                histogramChart.SetSourceData(worksheet.Range("F1:G8"));

                // 设置标题
                histogramChart.HasTitle = true;
                histogramChart.ChartTitle.Text = "销售额频率分布";

                // 设置图例
                histogramChart.HasLegend = false;

                // 设置坐标轴标题
                histogramChart.Axes(XlAxisType.xlCategory).HasTitle = true;
                histogramChart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "销售额区间";

                histogramChart.Axes(XlAxisType.xlValue).HasTitle = true;
                histogramChart.Axes(XlAxisType.xlValue).AxisTitle.Text = "频数";

                // 设置数字格式
                worksheet.Range("B2:B21").NumberFormat = "¥#,##0";
                worksheet.Range("H2:H8").NumberFormat = "0.00%";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"FrequencyDistribution_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示频率分布分析: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 频率分布分析时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据可视化示例
        /// 演示多种数据可视化技术
        /// </summary>
        static void DataVisualizationExample()
        {
            Console.WriteLine("=== 数据可视化示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数据可视化";

                // 创建产品销售数据
                worksheet.Range("A1").Value = "产品类别";
                worksheet.Range("B1").Value = "产品名称";
                worksheet.Range("C1").Value = "销售额";
                worksheet.Range("D1").Value = "市场份额";

                object[,] productData = {
                    {"电子产品", "笔记本电脑", 500000, 0.3},
                    {"电子产品", "手机", 300000, 0.18},
                    {"电子产品", "平板电脑", 200000, 0.12},
                    {"家居用品", "沙发", 150000, 0.09},
                    {"家居用品", "床", 120000, 0.07},
                    {"家居用品", "餐桌", 100000, 0.06},
                    {"服装", "T恤", 80000, 0.05},
                    {"服装", "牛仔裤", 70000, 0.03}
                };

                var dataRange = worksheet.Range("A2:D9");
                dataRange.Value = productData;

                // 创建柱形图
                var columnChartObject = worksheet.ChartObjects().Add(300, 50, 500, 350);
                var columnChart = columnChartObject.Chart;

                // 设置图表类型为柱形图
                columnChart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                columnChart.SetSourceData(worksheet.Range("A1:C9"));

                // 设置标题
                columnChart.HasTitle = true;
                columnChart.ChartTitle.Text = "产品销售额对比";

                // 设置图例
                columnChart.HasLegend = true;
                columnChart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 设置坐标轴标题
                columnChart.Axes(XlAxisType.xlCategory).HasTitle = true;
                columnChart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "产品";

                columnChart.Axes(XlAxisType.xlValue).HasTitle = true;
                columnChart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                // 创建饼图
                var pieChartObject = worksheet.ChartObjects().Add(850, 50, 400, 300);
                var pieChart = pieChartObject.Chart;

                // 设置图表类型为饼图
                pieChart.ChartType = MsoChartType.xlPie;

                // 设置数据源
                pieChart.SetSourceData(worksheet.Range("A1:A9,C1:C9"));

                // 设置标题
                pieChart.HasTitle = true;
                pieChart.ChartTitle.Text = "产品销售额占比";

                // 设置图例
                pieChart.HasLegend = true;
                pieChart.Legend.Position = XlLegendPosition.xlLegendPositionRight;

                // 设置数据标签
                var pieSeries = pieChart.SeriesCollection()[1];
                pieSeries.HasDataLabels = true;
                pieSeries.DataLabels().ShowCategoryName = true;
                pieSeries.DataLabels().ShowPercentage = true;

                // 创建组合图表（柱形图+折线图）
                var comboWorksheet = workbook.Worksheets.Add() as IExcelWorksheet;
                comboWorksheet.Name = "组合图表";

                // 复制数据到新工作表
                comboWorksheet.Range("A1").Value = "月份";
                comboWorksheet.Range("B1").Value = "销售额";
                comboWorksheet.Range("C1").Value = "增长率";

                object[,] monthlyData = {
                    {"1月", 100000, 0.05},
                    {"2月", 120000, 0.20},
                    {"3月", 140000, 0.167},
                    {"4月", 130000, -0.071},
                    {"5月", 150000, 0.154},
                    {"6月", 160000, 0.067},
                    {"7月", 170000, 0.0625},
                    {"8月", 180000, 0.0588},
                    {"9月", 190000, 0.0556},
                    {"10月", 200000, 0.0526},
                    {"11月", 210000, 0.05},
                    {"12月", 220000, 0.0476}
                };

                var comboDataRange = comboWorksheet.Range("A2:C13");
                comboDataRange.Value = monthlyData;

                // 创建组合图表对象
                var comboChartObject = comboWorksheet.ChartObjects().Add(300, 50, 600, 400);
                var comboChart = comboChartObject.Chart;

                // 设置图表类型为柱形图
                comboChart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                comboChart.SetSourceData(comboWorksheet.Range("A1:C13"));

                // 设置标题
                comboChart.HasTitle = true;
                comboChart.ChartTitle.Text = "销售额与增长率组合图表";

                // 设置图例
                comboChart.HasLegend = true;
                comboChart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 修改增长率系列为折线图
                var growthSeries = comboChart.SeriesCollection()[3];
                growthSeries.ChartType = MudTools.OfficeInterop.Excel.XlChartType.xlLine;

                // 设置次坐标轴
                growthSeries.AxisGroup = XlAxisGroup.xlSecondary;

                // 设置坐标轴标题
                comboChart.Axes(XlAxisType.xlCategory).HasTitle = true;
                comboChart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "月份";

                comboChart.Axes(XlAxisType.xlValue).HasTitle = true;
                comboChart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                comboChart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).HasTitle = true;
                comboChart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Text = "增长率";

                // 设置数字格式
                worksheet.Range("C2:C9").NumberFormat = "¥#,##0";
                worksheet.Range("D2:D9").NumberFormat = "0.00%";

                comboWorksheet.Range("B2:B13").NumberFormat = "¥#,##0";
                comboWorksheet.Range("C2:C13").NumberFormat = "0.00%";

                // 自动调整列宽
                worksheet.Columns.AutoFit();
                comboWorksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"DataVisualization_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数据可视化: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数据可视化时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 综合数据分析报告示例
        /// 演示如何创建综合的数据分析报告
        /// </summary>
        static void ComprehensiveAnalysisReportExample()
        {
            Console.WriteLine("=== 综合数据分析报告示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 创建数据源工作表
                var dataWorksheet = workbook.Worksheets.Add() as IExcelWorksheet;
                dataWorksheet.Name = "原始数据";

                // 创建销售数据
                dataWorksheet.Range("A1").Value = "日期";
                dataWorksheet.Range("B1").Value = "产品类别";
                dataWorksheet.Range("C1").Value = "产品名称";
                dataWorksheet.Range("D1").Value = "销售地区";
                dataWorksheet.Range("E1").Value = "销售人员";
                dataWorksheet.Range("F1").Value = "销售数量";
                dataWorksheet.Range("G1").Value = "单价";
                dataWorksheet.Range("H1").Value = "销售额";

                object[,] salesData = {
                    {"2023-01-01", "电子产品", "笔记本电脑", "北京", "张三", 2, 50000, 100000},
                    {"2023-01-02", "电子产品", "手机", "上海", "李四", 5, 6000, 30000},
                    {"2023-01-03", "家居用品", "沙发", "广州", "王五", 1, 150000, 150000},
                    {"2023-01-04", "服装", "T恤", "深圳", "赵六", 20, 50, 1000},
                    {"2023-01-05", "电子产品", "平板电脑", "北京", "张三", 3, 40000, 120000},
                    {"2023-01-06", "家居用品", "床", "上海", "李四", 1, 80000, 80000},
                    {"2023-01-07", "服装", "牛仔裤", "广州", "王五", 15, 100, 1500},
                    {"2023-01-08", "电子产品", "手机", "深圳", "赵六", 8, 6000, 48000},
                    {"2023-01-09", "家居用品", "餐桌", "北京", "张三", 1, 50000, 50000},
                    {"2023-01-10", "服装", "外套", "上海", "李四", 5, 300, 1500},
                    {"2023-02-01", "电子产品", "笔记本电脑", "广州", "王五", 1, 50000, 50000},
                    {"2023-02-02", "电子产品", "手机", "深圳", "赵六", 10, 6000, 60000},
                    {"2023-02-03", "家居用品", "沙发", "北京", "张三", 2, 150000, 300000},
                    {"2023-02-04", "服装", "连衣裙", "上海", "李四", 8, 200, 1600},
                    {"2023-02-05", "电子产品", "台式电脑", "广州", "王五", 2, 80000, 160000}
                };

                var dataRange = dataWorksheet.Range("A2:H16");
                dataRange.Value = salesData;

                // 创建汇总分析工作表
                var summaryWorksheet = workbook.Worksheets.Add() as IExcelWorksheet;
                summaryWorksheet.Name = "汇总分析";

                // 汇总统计
                summaryWorksheet.Range("A1").Value = "销售数据分析报告";
                summaryWorksheet.Range("A1").Font.Bold = true;
                summaryWorksheet.Range("A1").Font.Size = 16;
                summaryWorksheet.Range("A1").Interior.Color = Color.DarkBlue;
                summaryWorksheet.Range("A1").Font.Color = Color.White;
                var titleRange = summaryWorksheet.Range("A1:E1");
                titleRange.Merge();
                titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                summaryWorksheet.Range("A3").Value = "总体统计";
                summaryWorksheet.Range("A3").Font.Bold = true;
                summaryWorksheet.Range("A3").Interior.Color = Color.LightBlue;

                summaryWorksheet.Range("A4").Value = "总销售额";
                summaryWorksheet.Range("B4").Formula = "=SUM(原始数据!H2:H16)";

                summaryWorksheet.Range("A5").Value = "总销售数量";
                summaryWorksheet.Range("B5").Formula = "=SUM(原始数据!F2:F16)";

                summaryWorksheet.Range("A6").Value = "平均单价";
                summaryWorksheet.Range("B6").Formula = "=AVERAGE(原始数据!G2:G16)";

                summaryWorksheet.Range("A7").Value = "销售记录数";
                summaryWorksheet.Range("B7").Formula = "=COUNT(原始数据!A2:A16)";

                // 按产品类别汇总
                summaryWorksheet.Range("A9").Value = "按产品类别汇总";
                summaryWorksheet.Range("A9").Font.Bold = true;
                summaryWorksheet.Range("A9").Interior.Color = Color.LightGreen;

                summaryWorksheet.Range("A10").Value = "产品类别";
                summaryWorksheet.Range("B10").Value = "销售额";
                summaryWorksheet.Range("C10").Value = "占比";

                summaryWorksheet.Range("A11").Value = "电子产品";
                summaryWorksheet.Range("B11").Formula = "=SUMIF(原始数据!B2:B16,\"电子产品\",原始数据!H2:H16)";
                summaryWorksheet.Range("C11").Formula = "=B11/B4";

                summaryWorksheet.Range("A12").Value = "家居用品";
                summaryWorksheet.Range("B12").Formula = "=SUMIF(原始数据!B2:B16,\"家居用品\",原始数据!H2:H16)";
                summaryWorksheet.Range("C12").Formula = "=B12/B4";

                summaryWorksheet.Range("A13").Value = "服装";
                summaryWorksheet.Range("B13").Formula = "=SUMIF(原始数据!B2:B16,\"服装\",原始数据!H2:H16)";
                summaryWorksheet.Range("C13").Formula = "=B13/B4";

                // 按销售地区汇总
                summaryWorksheet.Range("A15").Value = "按销售地区汇总";
                summaryWorksheet.Range("A15").Font.Bold = true;
                summaryWorksheet.Range("A15").Interior.Color = Color.LightYellow;

                summaryWorksheet.Range("A16").Value = "销售地区";
                summaryWorksheet.Range("B16").Value = "销售额";
                summaryWorksheet.Range("C16").Value = "占比";

                summaryWorksheet.Range("A17").Value = "北京";
                summaryWorksheet.Range("B17").Formula = "=SUMIF(原始数据!D2:D16,\"北京\",原始数据!H2:H16)";
                summaryWorksheet.Range("C17").Formula = "=B17/B4";

                summaryWorksheet.Range("A18").Value = "上海";
                summaryWorksheet.Range("B18").Formula = "=SUMIF(原始数据!D2:D16,\"上海\",原始数据!H2:H16)";
                summaryWorksheet.Range("C18").Formula = "=B18/B4";

                summaryWorksheet.Range("A19").Value = "广州";
                summaryWorksheet.Range("B19").Formula = "=SUMIF(原始数据!D2:D16,\"广州\",原始数据!H2:H16)";
                summaryWorksheet.Range("C19").Formula = "=B19/B4";

                summaryWorksheet.Range("A20").Value = "深圳";
                summaryWorksheet.Range("B20").Formula = "=SUMIF(原始数据!D2:D16,\"深圳\",原始数据!H2:H16)";
                summaryWorksheet.Range("C20").Formula = "=B20/B4";

                // 创建图表
                // 产品类别销售额饼图
                var categoryChartObject = summaryWorksheet.ChartObjects().Add(300, 50, 400, 300);
                var categoryChart = categoryChartObject.Chart;

                categoryChart.ChartType = MsoChartType.xlPie;
                categoryChart.SetSourceData(summaryWorksheet.Range("A11:B13"));
                categoryChart.HasTitle = true;
                categoryChart.ChartTitle.Text = "按产品类别销售额分布";
                categoryChart.HasLegend = true;

                var categorySeries = categoryChart.SeriesCollection()[1];
                categorySeries.HasDataLabels = true;
                categorySeries.DataLabels().ShowCategoryName = true;
                categorySeries.DataLabels().ShowPercentage = true;

                // 销售地区销售额柱形图
                var regionChartObject = summaryWorksheet.ChartObjects().Add(300, 400, 500, 350);
                var regionChart = regionChartObject.Chart;

                regionChart.ChartType = MsoChartType.xlColumnClustered;
                regionChart.SetSourceData(summaryWorksheet.Range("A17:B20"));
                regionChart.HasTitle = true;
                regionChart.ChartTitle.Text = "按销售地区销售额对比";
                regionChart.HasLegend = false;

                regionChart.Axes(XlAxisType.xlCategory).HasTitle = true;
                regionChart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "销售地区";

                regionChart.Axes(XlAxisType.xlValue).HasTitle = true;
                regionChart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                // 设置数字格式
                dataWorksheet.Range("F2:F16").NumberFormat = "0";
                dataWorksheet.Range("G2:H16").NumberFormat = "¥#,##0";

                summaryWorksheet.Range("B4:B7").NumberFormat = "¥#,##0";
                summaryWorksheet.Range("B11:B20").NumberFormat = "¥#,##0";
                summaryWorksheet.Range("C11:C20").NumberFormat = "0.00%";

                // 自动调整列宽
                dataWorksheet.Columns.AutoFit();
                summaryWorksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ComprehensiveAnalysisReport_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示综合数据分析报告: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 综合数据分析报告时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}