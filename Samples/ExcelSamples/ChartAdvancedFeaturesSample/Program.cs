//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace ChartAdvancedFeaturesSample
{
    /// <summary>
    /// 图表高级功能示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel创建和配置具有高级功能的图表
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("图表高级功能示例");
            Console.WriteLine("================");
            Console.WriteLine();

            // 演示趋势线分析
            TrendlineAnalysisExample();

            // 演示误差线应用
            ErrorBarsExample();

            // 演示数据标签高级设置
            AdvancedDataLabelsExample();

            // 演示图表区域格式设置
            ChartAreaFormattingExample();

            // 演示图表样式和主题应用
            ChartStylesAndThemesExample();

            // 演示组合图表高级功能
            AdvancedCombinationChartExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 趋势线分析示例
        /// 演示如何在图表中添加和配置各种类型的趋势线
        /// </summary>
        static void TrendlineAnalysisExample()
        {
            Console.WriteLine("=== 趋势线分析示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "趋势线分析";

                // 创建销售数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";

                object[,] salesData = {
                    {"1月", 100},
                    {"2月", 120},
                    {"3月", 140},
                    {"4月", 130},
                    {"5月", 150},
                    {"6月", 160},
                    {"7月", 170},
                    {"8月", 180},
                    {"9月", 190},
                    {"10月", 200},
                    {"11月", 210},
                    {"12月", 220}
                };

                var dataRange = worksheet.Range("A2:B13");
                dataRange.Value = salesData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 600, 400);
                var chart = chartObject.Chart;

                // 设置图表类型为散点图
                chart.ChartType = MsoChartType.xlXYScatter;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "销售额趋势分析";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 添加线性趋势线
                var series = chart.SeriesCollection()[1];
                var linearTrendline = series.Trendlines().Add(XlTrendlineType.xlLinear);
                linearTrendline.DisplayEquation = true;
                linearTrendline.DisplayRSquared = true;
                linearTrendline.Name = "线性趋势";

                // 添加指数趋势线
                var exponentialTrendline = series.Trendlines().Add(XlTrendlineType.xlExponential);
                exponentialTrendline.DisplayEquation = true;
                exponentialTrendline.DisplayRSquared = true;
                exponentialTrendline.Name = "指数趋势";

                // 添加多项式趋势线
                var polynomialTrendline = series.Trendlines().Add(XlTrendlineType.xlPolynomial, 2);
                polynomialTrendline.DisplayEquation = true;
                polynomialTrendline.DisplayRSquared = true;
                polynomialTrendline.Name = "二次多项式趋势";

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "月份";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"TrendlineAnalysis_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示趋势线分析: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 趋势线分析时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 误差线应用示例
        /// 演示如何在图表中添加和配置误差线
        /// </summary>
        static void ErrorBarsExample()
        {
            Console.WriteLine("=== 误差线应用示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "误差线";

                // 创建实验数据
                worksheet.Range("A1").Value = "实验组";
                worksheet.Range("B1").Value = "测量值";
                worksheet.Range("C1").Value = "标准差";

                object[,] experimentData = {
                    {"实验1", 95.2, 3.1},
                    {"实验2", 87.6, 2.8},
                    {"实验3", 92.1, 4.2},
                    {"实验4", 89.8, 3.5},
                    {"实验5", 91.3, 2.9}
                };

                var dataRange = worksheet.Range("A2:C6");
                dataRange.Value = experimentData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 500, 350);
                var chart = chartObject.Chart;

                // 设置图表类型为柱形图
                chart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "实验测量结果与误差分析";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 添加误差线
                var series = chart.SeriesCollection()[1];
                var errorBars = series.ErrorBar(XlErrorBarDirection.xlY,
                    XlErrorBarInclude.xlErrorBarIncludeBoth,
                    XlErrorBarType.xlErrorBarTypeCustom,
                    dataRange.Columns[3], dataRange.Columns[3]);

                // 设置误差线格式
                errorBars.EndStyle = XlEndStyleCap.xlCap;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "实验组";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "测量值";

                // 设置数字格式
                worksheet.Range("B2:B6").NumberFormat = "0.0";
                worksheet.Range("C2:C6").NumberFormat = "0.0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ErrorBars_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示误差线应用: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 误差线应用时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据标签高级设置示例
        /// 演示如何配置高级数据标签选项
        /// </summary>
        static void AdvancedDataLabelsExample()
        {
            Console.WriteLine("=== 数据标签高级设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数据标签";

                // 创建市场份额数据
                worksheet.Range("A1").Value = "产品";
                worksheet.Range("B1").Value = "市场份额";

                object[,] marketData = {
                    {"产品A", 35},
                    {"产品B", 25},
                    {"产品C", 20},
                    {"产品D", 12},
                    {"产品E", 8}
                };

                var dataRange = worksheet.Range("A2:B6");
                dataRange.Value = marketData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 500, 400);
                var chart = chartObject.Chart;

                // 设置图表类型为饼图
                chart.ChartType = MsoChartType.xlPie;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "产品市场份额";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionRight;

                // 配置数据标签
                var series = chart.SeriesCollection()[1];
                series.HasDataLabels = true;

                var dataLabels = series.DataLabels();
                dataLabels.ShowCategoryName = true;
                dataLabels.ShowValue = true;
                dataLabels.ShowPercentage = true;

                // 设置数据标签位置
                dataLabels.Position = XlDataLabelPosition.xlLabelPositionOutsideEnd;

                // 设置数据标签格式
                dataLabels.Font.Bold = true;
                dataLabels.Font.Size = 10;

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"AdvancedDataLabels_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数据标签高级设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数据标签高级设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 图表区域格式设置示例
        /// 演示如何设置图表区域的高级格式
        /// </summary>
        static void ChartAreaFormattingExample()
        {
            Console.WriteLine("=== 图表区域格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "区域格式";

                // 创建销售数据
                worksheet.Range("A1").Value = "季度";
                worksheet.Range("B1").Value = "产品A";
                worksheet.Range("C1").Value = "产品B";
                worksheet.Range("D1").Value = "产品C";

                object[,] salesData = {
                    {"Q1", 100, 80, 120},
                    {"Q2", 120, 90, 130},
                    {"Q3", 140, 100, 110},
                    {"Q4", 130, 110, 140}
                };

                var dataRange = worksheet.Range("A2:D5");
                dataRange.Value = salesData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 600, 400);
                var chart = chartObject.Chart;

                // 设置图表类型为柱形图
                chart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "季度销售对比";
                chart.ChartTitle.Font.Size = 16;
                chart.ChartTitle.Font.Bold = true;

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 设置图表区格式
                chart.ChartArea.Format.Fill.ForeColor.RGB = Color.FromArgb(240, 240, 240).ToArgb();
                chart.ChartArea.Format.Fill.Transparency = 0.2f;

                // 设置绘图区格式
                chart.PlotArea.Format.Fill.ForeColor.RGB = Color.White.ToArgb();
                chart.PlotArea.Format.Line.Visible = true;
                chart.PlotArea.Format.Line.ForeColor.RGB = Color.Black.ToArgb();
                chart.PlotArea.Format.Line.Weight = 1;

                // 设置坐标轴格式
                var categoryAxis = chart.Axes(XlAxisType.xlCategory);
                categoryAxis.HasTitle = true;
                categoryAxis.AxisTitle.Text = "季度";
                categoryAxis.TickLabels.Font.Size = 10;

                var valueAxis = chart.Axes(XlAxisType.xlValue);
                valueAxis.HasTitle = true;
                valueAxis.AxisTitle.Text = "销售额";
                valueAxis.TickLabels.Font.Size = 10;
                valueAxis.MajorGridlines.Format.Line.Visible = true;
                valueAxis.MajorGridlines.Format.Line.ForeColor.RGB = Color.LightGray.ToArgb();

                // 设置数据系列格式
                var seriesCollection = chart.SeriesCollection();
                for (int i = 1; i <= seriesCollection.Count; i++)
                {
                    var series = seriesCollection[i];
                    series.Format.Line.Weight = 2;
                }

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ChartAreaFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示图表区域格式设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 图表区域格式设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 图表样式和主题应用示例
        /// 演示如何应用图表样式和主题
        /// </summary>
        static void ChartStylesAndThemesExample()
        {
            Console.WriteLine("=== 图表样式和主题应用示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "样式主题";

                // 创建数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "成本";
                worksheet.Range("D1").Value = "利润";

                object[,] financialData = {
                    {"1月", 100000, 70000, 30000},
                    {"2月", 120000, 80000, 40000},
                    {"3月", 140000, 90000, 50000},
                    {"4月", 130000, 85000, 45000},
                    {"5月", 150000, 95000, 55000},
                    {"6月", 160000, 100000, 60000}
                };

                var dataRange = worksheet.Range("A2:D7");
                dataRange.Value = financialData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 700, 450);
                var chart = chartObject.Chart;

                // 设置图表类型为折线图
                chart.ChartType = MsoChartType.xlLineMarkers;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "财务数据分析";
                chart.ChartTitle.Font.Size = 18;
                chart.ChartTitle.Font.Bold = true;

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 应用图表样式
                chart.ChartStyle = MsoChartType.xl3DLine; // 应用预定义样式

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "月份";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "金额";

                // 设置数字格式
                worksheet.Range("B2:D7").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ChartStylesAndThemes_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示图表样式和主题应用: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 图表样式和主题应用时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 组合图表高级功能示例
        /// 演示如何创建具有高级功能的组合图表
        /// </summary>
        static void AdvancedCombinationChartExample()
        {
            Console.WriteLine("=== 组合图表高级功能示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "组合图表";

                // 创建销售和增长率数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "增长率";

                object[,] salesGrowthData = {
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

                var dataRange = worksheet.Range("A2:C13");
                dataRange.Value = salesGrowthData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 700, 450);
                var chart = chartObject.Chart;

                // 设置图表类型为柱形图
                chart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "销售额与增长率组合分析";
                chart.ChartTitle.Font.Size = 16;
                chart.ChartTitle.Font.Bold = true;

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 修改增长率系列为折线图
                var growthSeries = chart.SeriesCollection()[3];
                growthSeries.ChartType = MsoChartType.xlLine;

                // 设置次坐标轴
                growthSeries.AxisGroup = XlAxisGroup.xlSecondary;

                // 为增长率系列添加数据标签
                growthSeries.HasDataLabels = true;
                growthSeries.DataLabels().NumberFormat = "0.00%";
                growthSeries.DataLabels().Font.Size = 9;

                // 为销售额系列添加趋势线
                var salesSeries = chart.SeriesCollection()[2];
                var trendline = salesSeries.Trendlines().Add(XlTrendlineType.xlLinear);
                trendline.DisplayEquation = true;
                trendline.DisplayRSquared = true;
                trendline.Name = "销售趋势";

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "月份";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).HasTitle = true;
                chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Text = "增长率";

                // 设置坐标轴格式
                chart.Axes(XlAxisType.xlValue).TickLabels.NumberFormat = "¥#,##0";
                chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).TickLabels.NumberFormat = "0.00%";

                // 设置图表区和绘图区格式
                chart.ChartArea.Format.Fill.ForeColor.RGB = Color.White.ToArgb();
                chart.PlotArea.Format.Fill.ForeColor.RGB = Color.FromArgb(250, 250, 250).ToArgb();

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"AdvancedCombinationChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示组合图表高级功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 组合图表高级功能时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}