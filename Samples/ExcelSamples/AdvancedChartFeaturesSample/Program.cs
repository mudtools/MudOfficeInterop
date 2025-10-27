//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace AdvancedChartFeaturesSample
{
    /// <summary>
    /// 高级图表功能示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel创建和配置高级图表功能
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("高级图表功能示例");
            Console.WriteLine("==============");
            Console.WriteLine();

            // 演示组合图表技术
            CombinationChartExample();

            // 演示动态图表效果
            DynamicChartExample();

            // 演示3D图表功能
            ThreeDChartExample();

            // 演示图表交互功能
            InteractiveChartExample();

            // 演示高级图表样式
            AdvancedChartStylingExample();

            // 演示图表模板应用
            ChartTemplateExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 组合图表技术示例
        /// 演示如何创建和配置组合图表
        /// </summary>
        static void CombinationChartExample()
        {
            Console.WriteLine("=== 组合图表技术示例 ===");

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
                worksheet.Range("C1").Value = "利润率";

                object[,] salesProfitData = {
                    {"1月", 100000, 0.2},
                    {"2月", 120000, 0.25},
                    {"3月", 140000, 0.22},
                    {"4月", 130000, 0.24},
                    {"5月", 150000, 0.26},
                    {"6月", 160000, 0.28}
                };

                var dataRange = worksheet.Range("A2:C7");
                dataRange.Value = salesProfitData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 600, 400);
                var chart = chartObject.Chart;

                // 设置图表类型为柱形图
                chart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "销售与利润率组合图表";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 修改利润率系列为折线图
                var profitSeries = chart.SeriesCollection()[3];
                profitSeries.ChartType = MsoChartType.xlLine;

                // 设置次坐标轴
                profitSeries.AxisGroup = XlAxisGroup.xlSecondary;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "月份";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).HasTitle = true;
                chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Text = "利润率";

                // 设置数字格式
                worksheet.Range("B2:B7").NumberFormat = "¥#,##0";
                worksheet.Range("C2:C7").NumberFormat = "0.00%";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = Path.Combine(AppContext.BaseDirectory, $"CombinationChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示组合图表技术: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 组合图表技术时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 动态图表效果示例
        /// 演示如何创建具有动态效果的图表
        /// </summary>
        static void DynamicChartExample()
        {
            Console.WriteLine("=== 动态图表效果示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "动态图表";

                // 创建季度销售数据
                worksheet.Range("A1").Value = "产品";
                worksheet.Range("B1").Value = "Q1";
                worksheet.Range("C1").Value = "Q2";
                worksheet.Range("D1").Value = "Q3";
                worksheet.Range("E1").Value = "Q4";

                object[,] quarterlyData = {
                    {"产品A", 100, 120, 140, 130},
                    {"产品B", 80, 90, 100, 110},
                    {"产品C", 120, 130, 110, 140}
                };

                var dataRange = worksheet.Range("A2:E4");
                dataRange.Value = quarterlyData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 600, 400);
                var chart = chartObject.Chart;

                // 设置图表类型为柱形图
                chart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "产品季度销售动态图表";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 设置动画效果
                chart.ApplyDataLabels(); // 应用数据类型

                // 设置数据标签
                var seriesCollection = chart.SeriesCollection();
                for (int i = 1; i <= seriesCollection.Count; i++)
                {
                    var series = seriesCollection[i];
                    series.HasDataLabels = true;
                    series.DataLabels().Position = XlDataLabelPosition.xlLabelPositionOutsideEnd;
                }

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "产品";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = Path.Combine(AppContext.BaseDirectory, $"DynamicChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示动态图表效果: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 动态图表效果时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 3D图表功能示例
        /// 演示如何创建和配置3D图表
        /// </summary>
        static void ThreeDChartExample()
        {
            Console.WriteLine("=== 3D图表功能示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "3D图表";

                // 创建地区销售数据
                worksheet.Range("A1").Value = "地区";
                worksheet.Range("B1").Value = "产品A";
                worksheet.Range("C1").Value = "产品B";
                worksheet.Range("D1").Value = "产品C";

                object[,] regionData = {
                    {"华东", 100, 80, 120},
                    {"华南", 120, 90, 130},
                    {"华北", 140, 100, 110},
                    {"华中", 130, 110, 140}
                };

                var dataRange = worksheet.Range("A2:D5");
                dataRange.Value = regionData;

                // 创建图表对象
                var chartObject = worksheet.ChartObjects().Add(300, 50, 600, 450);
                var chart = chartObject.Chart;

                // 设置图表类型为3D柱形图
                chart.ChartType = MsoChartType.xl3DColumn;

                // 设置数据源
                chart.SetSourceData(dataRange);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "各地区产品销售3D图表";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 设置3D视图
                chart.Rotation = 20;  // 旋转角度
                chart.Elevation = 15; // 仰角
                chart.DepthPercent = 100; // 深度百分比
                chart.GapDepth = 150; // 间距深度

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "地区";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = Path.Combine(AppContext.BaseDirectory, $"ThreeDChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示3D图表功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 3D图表功能时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 图表交互功能示例
        /// 演示如何创建具有交互功能的图表
        /// </summary>
        static void InteractiveChartExample()
        {
            Console.WriteLine("=== 图表交互功能示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "交互图表";

                // 创建详细销售数据
                worksheet.Range("A1").Value = "日期";
                worksheet.Range("B1").Value = "产品类别";
                worksheet.Range("C1").Value = "销售额";

                object[,] salesData = {
                    {"2023-01-01", "电子产品", 10000},
                    {"2023-01-02", "家居用品", 8000},
                    {"2023-01-03", "服装", 5000},
                    {"2023-01-04", "电子产品", 12000},
                    {"2023-01-05", "家居用品", 9000},
                    {"2023-01-06", "服装", 6000},
                    {"2023-01-07", "电子产品", 15000},
                    {"2023-01-08", "家居用品", 10000},
                    {"2023-01-09", "服装", 7000}
                };

                var dataRange = worksheet.Range("A2:C10");
                dataRange.Value = salesData;

                // 创建数据透视表
                var pivotCache = worksheet.PivotCaches().Create(XlPivotTableSourceType.xlPivotTable, dataRange);
                var pivotTable = worksheet.PivotTables().Add(pivotCache, worksheet.Range("E1"), "SalesPivotTable");

                // 配置数据透视表字段
                pivotTable.PivotFields("产品类别").Orientation = XlPivotFieldOrientation.xlRowField;
                pivotTable.PivotFields("日期").Orientation = XlPivotFieldOrientation.xlColumnField;
                pivotTable.PivotFields("销售额").Orientation = XlPivotFieldOrientation.xlDataField;

                // 创建基于数据透视表的图表
                var chartObject = worksheet.ChartObjects().Add(300, 200, 600, 400);
                var chart = chartObject.Chart;

                // 设置图表类型为柱形图
                chart.ChartType = MsoChartType.xlColumnClustered;

                // 设置数据源为数据透视表
                chart.SetSourceData(pivotTable.TableRange1);

                // 设置标题
                chart.HasTitle = true;
                chart.ChartTitle.Text = "交互式销售分析图表";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "产品类别";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "销售额";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = Path.Combine(AppContext.BaseDirectory, $"InteractiveChart_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示图表交互功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 图表交互功能时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 高级图表样式示例
        /// 演示如何应用高级图表样式和格式
        /// </summary>
        static void AdvancedChartStylingExample()
        {
            Console.WriteLine("=== 高级图表样式示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "高级样式";

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
                chart.ChartTitle.Text = "产品市场份额分析";
                chart.ChartTitle.Font.Size = 16;
                chart.ChartTitle.Font.Bold = true;

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionRight;

                // 应用图表样式
                chart.ChartStyle = MsoChartType.xl3DColumn; // 应用预定义样式

                // 设置数据标签
                var series = chart.SeriesCollection()[1];
                series.HasDataLabels = true;

                var dataLabels = series.DataLabels();
                dataLabels.ShowCategoryName = true;
                dataLabels.ShowValue = false;
                dataLabels.ShowPercentage = true;
                dataLabels.Position = XlDataLabelPosition.xlLabelPositionBestFit;

                // 设置数据系列格式
                series.Format.Fill.Solid();
                series.Format.Fill.ForeColor.RGB = Color.Blue.ToArgb();

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = Path.Combine(AppContext.BaseDirectory, $"AdvancedChartStyling_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示高级图表样式: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 高级图表样式时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 图表模板应用示例
        /// 演示如何使用和应用图表模板
        /// </summary>
        static void ChartTemplateExample()
        {
            Console.WriteLine("=== 图表模板应用示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "模板应用";

                // 创建销售数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "目标";

                object[,] salesData = {
                    {"1月", 100000, 120000},
                    {"2月", 120000, 125000},
                    {"3月", 140000, 130000},
                    {"4月", 130000, 135000},
                    {"5月", 150000, 140000},
                    {"6月", 160000, 145000}
                };

                var dataRange = worksheet.Range("A2:C7");
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
                chart.ChartTitle.Text = "销售目标达成情况";

                // 设置图例
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

                // 应用图表布局
                chart.ApplyLayout(3); // 应用预定义布局

                // 应用图表样式
                chart.ChartStyle = MsoChartType.xlBarStacked; // 应用预定义样式

                // 设置坐标轴标题
                chart.Axes(XlAxisType.xlCategory).HasTitle = true;
                chart.Axes(XlAxisType.xlCategory).AxisTitle.Text = "月份";

                chart.Axes(XlAxisType.xlValue).HasTitle = true;
                chart.Axes(XlAxisType.xlValue).AxisTitle.Text = "金额";

                // 设置数字格式
                worksheet.Range("B2:C7").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = Path.Combine(AppContext.BaseDirectory, $"ChartTemplate_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示图表模板应用: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 图表模板应用时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}