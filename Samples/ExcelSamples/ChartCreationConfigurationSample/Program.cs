//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;

namespace ChartCreationConfigurationSample
{
    /// <summary>
    /// 图表创建与配置示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel创建和配置各种类型的图表
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("图表创建与配置示例");
            Console.WriteLine("==================");
            Console.WriteLine();

            // 演示基础柱形图创建
            BasicColumnChartExample();

            // 演示折线图创建
            LineChartExample();

            // 演示饼图创建
            PieChartExample();

            // 演示条形图创建
            BarChartExample();

            // 演示面积图创建
            AreaChartExample();

            // 演示散点图创建
            ScatterChartExample();

            // 演示组合图表创建
            CombinationChartExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 基础柱形图创建示例
        /// 演示如何创建和配置基础柱形图
        /// </summary>
        static void BasicColumnChartExample()
        {
            Console.WriteLine("=== 基础柱形图创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "柱形图";

                // 创建销售数据
                worksheet.Range("A1").Value = "产品";
                worksheet.Range("B1").Value = "销量";

                object[,] salesData = {
                    {"产品A", 100},
                    {"产品B", 150},
                    {"产品C", 120},
                    {"产品D", 180},
                    {"产品E", 90}
                };

                var dataRange = worksheet.Range("A2:B6");
                dataRange.Value = salesData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 400, 300);
                var chart = chartObject.Chart;

                // 设置图表类型为簇状柱形图
                chart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "产品销量柱形图";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "产品";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销量";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"BasicColumnChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建基础柱形图: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建基础柱形图时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 折线图创建示例
        /// 演示如何创建和配置折线图
        /// </summary>
        static void LineChartExample()
        {
            Console.WriteLine("=== 折线图创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "折线图";

                // 创建月度销售数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "产品A";
                worksheet.Range("C1").Value = "产品B";
                worksheet.Range("D1").Value = "产品C";

                object[,] monthlyData = {
                    {"1月", 100, 80, 120},
                    {"2月", 120, 90, 130},
                    {"3月", 140, 100, 110},
                    {"4月", 130, 110, 140},
                    {"5月", 150, 120, 150},
                    {"6月", 160, 130, 160}
                };

                var dataRange = worksheet.Range("A2:D7");
                dataRange.Value = monthlyData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 500, 350);
                var chart = chartObject.Chart;

                // 设置图表类型为折线图
                chart.ChartType = MsoChartType.xlLine;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "月度销售趋势";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "月份";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销量";

                // 设置数据标签
                var seriesCollection = chart.SeriesCollection();
                for (int i = 1; i <= seriesCollection.Count; i++)
                {
                    var series = seriesCollection[i];
                    series.HasDataLabels = true;
                    series.DataLabels().Position = XlDataLabelPosition.xlLabelPositionAbove;
                }

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"LineChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建折线图: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建折线图时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 饼图创建示例
        /// 演示如何创建和配置饼图
        /// </summary>
        static void PieChartExample()
        {
            Console.WriteLine("=== 饼图创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "饼图";

                // 创建市场份额数据
                worksheet.Range("A1").Value = "产品";
                worksheet.Range("B1").Value = "市场份额";

                object[,] marketData = {
                    {"产品A", 30},
                    {"产品B", 25},
                    {"产品C", 20},
                    {"产品D", 15},
                    {"产品E", 10}
                };

                var dataRange = worksheet.Range("A2:B6");
                dataRange.Value = marketData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 400, 300);
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

                // 设置数据标签
                var series = chart.SeriesCollection()[1];
                series.HasDataLabels = true;
                series.DataLabels().ShowCategoryName = true;
                series.DataLabels().ShowValue = false;
                series.DataLabels().ShowPercentage = true;

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"PieChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建饼图: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建饼图时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 条形图创建示例
        /// 演示如何创建和配置条形图
        /// </summary>
        static void BarChartExample()
        {
            Console.WriteLine("=== 条形图创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "条形图";

                // 创建地区销售数据
                worksheet.Range("A1").Value = "地区";
                worksheet.Range("B1").Value = "销售额";

                object[,] regionData = {
                    {"华东", 1200000},
                    {"华南", 1000000},
                    {"华北", 900000},
                    {"华中", 800000},
                    {"西南", 700000},
                    {"东北", 600000}
                };

                var dataRange = worksheet.Range("A2:B7");
                dataRange.Value = regionData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 500, 350);
                var chart = chartObject.Chart;

                // 设置图表类型为条形图
                chart.ChartType = MsoChartType.xlBarClustered;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "各地区销售额";

                // 设置图例
                chart.HasLegend = false;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "地区";

                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "销售额";

                // 设置数据标签
                var series = chart.SeriesCollection()[1];
                series.HasDataLabels = true;
                series.DataLabels().Position = XlDataLabelPosition.xlLabelPositionInsideEnd;
                series.DataLabels().NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"BarChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建条形图: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建条形图时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 面积图创建示例
        /// 演示如何创建和配置面积图
        /// </summary>
        static void AreaChartExample()
        {
            Console.WriteLine("=== 面积图创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "面积图";

                // 创建季度数据
                worksheet.Range("A1").Value = "季度";
                worksheet.Range("B1").Value = "产品A";
                worksheet.Range("C1").Value = "产品B";
                worksheet.Range("D1").Value = "产品C";

                object[,] quarterlyData = {
                    {"Q1", 100, 80, 120},
                    {"Q2", 120, 90, 130},
                    {"Q3", 140, 100, 110},
                    {"Q4", 130, 110, 140}
                };

                var dataRange = worksheet.Range("A2:D5");
                dataRange.Value = quarterlyData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 500, 350);
                var chart = chartObject.Chart;

                // 设置图表类型为面积图
                chart.ChartType = MsoChartType.xlArea;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "季度销售面积图";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "季度";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销量";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"AreaChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建面积图: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建面积图时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 散点图创建示例
        /// 演示如何创建和配置散点图
        /// </summary>
        static void ScatterChartExample()
        {
            Console.WriteLine("=== 散点图创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "散点图";

                // 创建身高体重数据
                worksheet.Range("A1").Value = "身高(cm)";
                worksheet.Range("B1").Value = "体重(kg)";

                object[,] bodyData = {
                    {160, 50},
                    {165, 55},
                    {170, 60},
                    {175, 65},
                    {180, 70},
                    {185, 75},
                    {162, 52},
                    {168, 58},
                    {173, 63},
                    {178, 68},
                    {183, 73}
                };

                var dataRange = worksheet.Range("A2:B12");
                dataRange.Value = bodyData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 500, 350);
                var chart = chartObject.Chart;

                // 设置图表类型为散点图
                chart.ChartType = MsoChartType.xlXYScatter;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "身高与体重关系";

                // 设置图例
                chart.HasLegend = false;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "身高(cm)";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "体重(kg)";

                // 添加趋势线
                var series = chart.SeriesCollection()[1];
                series.Trendlines().Add(XlTrendlineType.xlLinear);

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ScatterChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建散点图: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建散点图时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 组合图表创建示例
        /// 演示如何创建和配置组合图表
        /// </summary>
        static void CombinationChartExample()
        {
            Console.WriteLine("=== 组合图表创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "组合图表";

                // 创建销售和利润数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "利润";
                worksheet.Range("D1").Value = "利润率";

                object[,] salesProfitData = {
                    {"1月", 100000, 20000, 0.2},
                    {"2月", 120000, 25000, 0.208},
                    {"3月", 140000, 28000, 0.2},
                    {"4月", 130000, 26000, 0.2},
                    {"5月", 150000, 30000, 0.2},
                    {"6月", 160000, 32000, 0.2}
                };

                var dataRange = worksheet.Range("A2:D7");
                dataRange.Value = salesProfitData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 600, 400);
                var chart = chartObject.Chart;

                // 设置图表类型为组合图（柱形图和折线图）
                chart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "销售与利润组合图表";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "月份";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "金额";

                // 修改第三个系列为折线图（利润率）
                var series3 = chart.SeriesCollection()[3];
                series3.ChartType = MsoChartType.xlLine;

                // 设置次坐标轴
                series3.AxisGroup = XlAxisGroup.xlSecondary;

                // 设置次坐标轴标题
                chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).HasTitle = true;
                chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Text = "利润率";

                // 设置数据标签
                var seriesCollection = chart.SeriesCollection();
                for (int i = 1; i <= seriesCollection.Count; i++)
                {
                    var series = seriesCollection[i];
                    if (i != 3) // 利润率系列不显示数据标签
                    {
                        series.HasDataLabels = true;
                    }
                }

                // 设置数字格式
                worksheet.Range("B2:C7").NumberFormat = "¥#,##0";
                worksheet.Range("D2:D7").NumberFormat = "0.00%";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"CombinationChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建组合图表: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建组合图表时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}