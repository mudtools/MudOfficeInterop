//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace SalesReportPivotTableSample
{
    /// <summary>
    /// 销售数据分析报表生成示例
    /// 演示如何使用MudTools.OfficeInterop.Excel创建复杂的多维度数据透视表报表
    /// 对应博文：基于.NET操作Excel COM组件生成数据透视报表
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  销售数据分析报表生成器");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("本示例将演示：");
            Console.WriteLine("1. 创建Excel应用程序和工作簿");
            Console.WriteLine("2. 生成销售数据源");
            Console.WriteLine("3. 创建多个维度的数据透视表");
            Console.WriteLine("4. 应用样式和格式");
            Console.WriteLine("5. 保存报表文件");
            Console.WriteLine();
            Console.WriteLine("开始生成报表...");
            Console.WriteLine();

            try
            {
                // ========== 1. 初始化Excel应用 ==========
                using var excelApp = ExcelFactory.BlankWorkbook();
                var workbook = excelApp.ActiveWorkbook;
                excelApp.Visible = true;

                Console.WriteLine("✓ Excel应用程序已初始化");

                // ========== 2. 导入/生成源数据 ==========
                Console.WriteLine("正在生成销售数据...");
                using var sourceWorksheet = workbook.ActiveSheetWrap;
                sourceWorksheet.Name = "源数据";

                // 创建表头
                sourceWorksheet.Range("A1").Value = "日期";
                sourceWorksheet.Range("B1").Value = "产品类别";
                sourceWorksheet.Range("C1").Value = "产品名称";
                sourceWorksheet.Range("D1").Value = "销售地区";
                sourceWorksheet.Range("E1").Value = "销售人员";
                sourceWorksheet.Range("F1").Value = "销售数量";
                sourceWorksheet.Range("G1").Value = "单价";
                sourceWorksheet.Range("H1").Value = "销售金额";

                // 格式化表头
                using var headerRange = sourceWorksheet.Range("A1:H1");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.Blue; // 蓝色背景
                headerRange.Font.Color = Color.White; // 白色文字

                // 准备销售数据（32条记录，覆盖2023年全年）
                object[,] salesData = {
                    {"2023-01-05", "电子产品", "笔记本电脑", "北京", "张三", 2, 5000, 10000},
                    {"2023-01-08", "电子产品", "手机", "北京", "张三", 3, 2000, 6000},
                    {"2023-01-10", "电子产品", "平板电脑", "北京", "张三", 2, 1500, 3000},
                    {"2023-01-12", "家居用品", "沙发", "北京", "张三", 1, 3000, 3000},
                    {"2023-01-15", "服装", "T恤", "北京", "张三", 10, 50, 500},
                    {"2023-02-05", "电子产品", "笔记本电脑", "上海", "李四", 1, 5000, 5000},
                    {"2023-02-08", "电子产品", "手机", "上海", "李四", 2, 2000, 4000},
                    {"2023-02-10", "家居用品", "床", "上海", "李四", 1, 2000, 2000},
                    {"2023-02-15", "服装", "牛仔裤", "上海", "李四", 5, 100, 500},
                    {"2023-03-05", "电子产品", "笔记本电脑", "广州", "王五", 2, 5000, 10000},
                    {"2023-03-08", "电子产品", "台式电脑", "广州", "王五", 2, 4000, 8000},
                    {"2023-03-12", "家居用品", "餐桌", "广州", "王五", 1, 1000, 1000},
                    {"2023-03-15", "服装", "外套", "广州", "王五", 3, 300, 900},
                    {"2023-04-05", "电子产品", "手机", "深圳", "赵六", 3, 2000, 6000},
                    {"2023-04-10", "电子产品", "平板电脑", "深圳", "赵六", 2, 1500, 3000},
                    {"2023-04-15", "家居用品", "衣柜", "深圳", "赵六", 1, 3000, 3000},
                    {"2023-04-20", "服装", "连衣裙", "深圳", "赵六", 4, 200, 800},
                    {"2023-05-08", "电子产品", "笔记本电脑", "北京", "张三", 1, 5000, 5000},
                    {"2023-05-12", "家居用品", "沙发", "北京", "张三", 2, 3000, 6000},
                    {"2023-06-05", "电子产品", "手机", "上海", "李四", 4, 2000, 8000},
                    {"2023-06-10", "服装", "T恤", "上海", "李四", 15, 50, 750},
                    {"2023-07-08", "电子产品", "台式电脑", "广州", "王五", 1, 4000, 4000},
                    {"2023-07-15", "家居用品", "床", "广州", "王五", 2, 2000, 4000},
                    {"2023-08-05", "电子产品", "笔记本电脑", "深圳", "赵六", 2, 5000, 10000},
                    {"2023-08-10", "服装", "外套", "深圳", "赵六", 5, 300, 1500},
                    {"2023-09-08", "电子产品", "平板电脑", "北京", "张三", 3, 1500, 4500},
                    {"2023-09-15", "家居用品", "餐桌", "北京", "张三", 2, 1000, 2000},
                    {"2023-10-05", "电子产品", "手机", "上海", "李四", 2, 2000, 4000},
                    {"2023-10-12", "服装", "牛仔裤", "上海", "李四", 8, 100, 800},
                    {"2023-11-08", "电子产品", "笔记本电脑", "广州", "王五", 3, 5000, 15000},
                    {"2023-11-15", "家居用品", "沙发", "广州", "王五", 1, 3000, 3000},
                    {"2023-12-05", "电子产品", "台式电脑", "深圳", "赵六", 2, 4000, 8000},
                    {"2023-12-10", "服装", "连衣裙", "深圳", "赵六", 6, 200, 1200}
                };

                // 写入数据
                using var dataRange = sourceWorksheet.Range("A2:H33");
                dataRange.Value = salesData;

                // 设置数字格式
                sourceWorksheet.Range("F2:F33").NumberFormat = "#,##0";
                sourceWorksheet.Range("G2:G33").NumberFormat = "#,##0.00";
                sourceWorksheet.Range("H2:H33").NumberFormat = "#,##0.00";

                // 自动调整列宽
                sourceWorksheet.Columns.AutoFit();

                Console.WriteLine("✓ 源数据已生成（32条销售记录）");

                // ========== 3. 创建产品销售透视表 ==========
                Console.WriteLine("正在创建产品销售透视表...");
                using var productPivotSheet = workbook.Worksheets.Add() as IExcelWorksheet;
                productPivotSheet.Name = "产品销售分析";

                using var pivotCache = workbook.PivotCaches().Create(
                    XlPivotTableSourceType.xlDatabase,
                    sourceWorksheet.Range("A1:H33").Address(external: true)
                );

                using var productPivot = productPivotSheet.PivotTables().Add(
                    pivotCache,
                    productPivotSheet.Range("A1"),
                    "产品销售透视表"
                );

                // 配置字段
                using var categoryField = productPivot.PivotFields("产品类别");
                categoryField.Orientation = XlPivotFieldOrientation.xlRowField;
                categoryField.Position = 1;

                using var productNameField = productPivot.PivotFields("产品名称");
                productNameField.Orientation = XlPivotFieldOrientation.xlRowField;
                productNameField.Position = 2;

                using var regionField = productPivot.PivotFields("销售地区");
                regionField.Orientation = XlPivotFieldOrientation.xlColumnField;

                // 添加销售金额汇总
                var amountField = productPivot.PivotFields("销售金额");
                amountField.Orientation = XlPivotFieldOrientation.xlDataField;
                amountField.Function = XlConsolidationFunction.xlSum;
                amountField.Name = "销售金额";
                amountField.NumberFormat = "#,##0.00";

                // 添加销售数量统计
                var qtyField = productPivot.PivotFields("销售数量");
                qtyField.Orientation = XlPivotFieldOrientation.xlDataField;
                qtyField.Function = XlConsolidationFunction.xlSum;
                qtyField.Name = "销售数量";
                qtyField.NumberFormat = "#,##0";

                // 添加平均单价
                var avgPriceField = productPivot.PivotFields("单价");
                avgPriceField.Orientation = XlPivotFieldOrientation.xlDataField;
                avgPriceField.Function = XlConsolidationFunction.xlAverage;
                avgPriceField.Name = "平均单价";
                avgPriceField.NumberFormat = "#,##0.00";

                // 设置透视表选项
                productPivot.RowGrand = true;
                productPivot.ColumnGrand = true;
                productPivot.HasAutoFormat = true;

                // 应用样式
                productPivot.TableStyle = "PivotStyleMedium9";
                productPivot.ShowTableStyleRowStripes = true;
                productPivot.ShowTableStyleColumnStripes = true;

                productPivotSheet.Columns.AutoFit();

                Console.WriteLine("✓ 产品销售透视表已创建");

                // ========== 4. 创建地区销售透视表 ==========
                Console.WriteLine("正在创建地区销售透视表...");
                using var regionPivotSheet = workbook.Worksheets.Add() as IExcelWorksheet;
                regionPivotSheet.Name = "地区销售分析";

                var regionPivot = regionPivotSheet.PivotTables().Add(
                    pivotCache,
                    regionPivotSheet.Range("A1"),
                    "地区销售透视表"
                );

                // 配置字段
                using var regionRowField = regionPivot.PivotFields("销售地区");
                regionRowField.Orientation = XlPivotFieldOrientation.xlRowField;

                using var regionColField = regionPivot.PivotFields("产品类别");
                regionColField.Orientation = XlPivotFieldOrientation.xlColumnField;

                // 添加销售金额和数量
                var regionAmountField = regionPivot.PivotFields("销售金额");
                regionAmountField.Orientation = XlPivotFieldOrientation.xlDataField;
                regionAmountField.Function = XlConsolidationFunction.xlSum;
                regionAmountField.Name = "销售金额";
                regionAmountField.NumberFormat = "#,##0.00";

                var regionQtyField = regionPivot.PivotFields("销售数量");
                regionQtyField.Orientation = XlPivotFieldOrientation.xlDataField;
                regionQtyField.Function = XlConsolidationFunction.xlSum;
                regionQtyField.Name = "销售数量";

                // 按销售金额降序排列
                regionPivot.PivotFields("销售地区").AutoSort(2, "销售金额");

                regionPivot.RowGrand = true;
                regionPivot.ColumnGrand = true;
                regionPivot.TableStyle = "PivotStyleMedium14";
                regionPivot.ShowTableStyleRowStripes = true;

                regionPivotSheet.Columns.AutoFit();

                Console.WriteLine("✓ 地区销售透视表已创建");

                // ========== 5. 创建销售人员业绩透视表 ==========
                Console.WriteLine("正在创建销售人员业绩透视表...");
                using var salespersonPivotSheet = workbook.Worksheets.Add() as IExcelWorksheet;
                salespersonPivotSheet.Name = "销售人员业绩";

                using var salespersonPivot = salespersonPivotSheet.PivotTables().Add(
                    pivotCache,
                    salespersonPivotSheet.Range("A1"),
                    "销售人员业绩透视表"
                );

                // 配置字段
                using var salespersonRowField = salespersonPivot.PivotFields("销售人员");
                salespersonRowField.Orientation = XlPivotFieldOrientation.xlRowField;

                using var salespersonColField = salespersonPivot.PivotFields("产品类别");
                salespersonColField.Orientation = XlPivotFieldOrientation.xlColumnField;

                var spAmountField = salespersonPivot.PivotFields("销售金额");
                spAmountField.Orientation = XlPivotFieldOrientation.xlDataField;
                spAmountField.Function = XlConsolidationFunction.xlSum;
                spAmountField.Name = "销售金额";
                spAmountField.NumberFormat = "#,##0.00";

                var spOrderCountField = salespersonPivot.PivotFields("销售数量");
                spOrderCountField.Orientation = XlPivotFieldOrientation.xlDataField;
                spOrderCountField.Function = XlConsolidationFunction.xlSum;
                spOrderCountField.Name = "订单数量";

                // 按销售金额降序排列
                salespersonPivot.PivotFields("销售人员").AutoSort(2, "销售金额");

                salespersonPivot.RowGrand = true;
                salespersonPivot.ColumnGrand = true;
                salespersonPivot.TableStyle = "PivotStyleMedium19";

                salespersonPivotSheet.Columns.AutoFit();

                Console.WriteLine("✓ 销售人员业绩透视表已创建");

                // ========== 6. 创建月度销售趋势透视表 ==========
                Console.WriteLine("正在创建月度销售趋势透视表...");
                using var monthlyPivotSheet = workbook.Worksheets.Add() as IExcelWorksheet;
                monthlyPivotSheet.Name = "月度销售趋势";

                // 添加月度列到源数据
                sourceWorksheet.Range("I1").Value = "月份";
                for (int row = 2; row <= 33; row++)
                {
                    var dateValue = DateTime.Parse(sourceWorksheet.Cells[row, 1].Value.ToString());
                    sourceWorksheet.Cells[row, 9].Value = dateValue.ToString("yyyy-MM");
                }

                // 重新创建数据透视表缓存（包含新列）
                using var monthlyPivotCache = workbook.PivotCaches().Create(
                    XlPivotTableSourceType.xlDatabase,
                    sourceWorksheet.Range("A1:I33").Address(external: true)
                );

                using var monthlyPivot = monthlyPivotSheet.PivotTables().Add(
                    monthlyPivotCache,
                    monthlyPivotSheet.Range("A1"),
                    "月度趋势透视表"
                );

                // 配置字段
                using var monthlyRowField = monthlyPivot.PivotFields("月份");
                monthlyRowField.Orientation = XlPivotFieldOrientation.xlRowField;

                using var monthlyColField = monthlyPivot.PivotFields("产品类别");
                monthlyColField.Orientation = XlPivotFieldOrientation.xlColumnField;

                var monthlyAmountField = monthlyPivot.PivotFields("销售金额");
                monthlyAmountField.Orientation = XlPivotFieldOrientation.xlDataField;
                monthlyAmountField.Function = XlConsolidationFunction.xlSum;
                monthlyAmountField.Name = "月度销售额";
                monthlyAmountField.NumberFormat = "#,##0.00";

                monthlyPivot.RowGrand = true;
                monthlyPivot.ColumnGrand = true;
                monthlyPivot.TableStyle = "PivotStyleMedium7";
                monthlyPivot.ShowTableStyleRowStripes = true;

                monthlyPivotSheet.Columns.AutoFit();

                Console.WriteLine("✓ 月度销售趋势透视表已创建");

                // ========== 7. 美化与格式化 ==========
                Console.WriteLine("正在美化报表格式...");
                // 为每个工作表添加标题
                AddReportTitle(productPivotSheet, "产品销售分析报表");
                AddReportTitle(regionPivotSheet, "地区销售分析报表");
                AddReportTitle(salespersonPivotSheet, "销售人员业绩报表");
                AddReportTitle(monthlyPivotSheet, "月度销售趋势报表");

                Console.WriteLine("✓ 报表格式已美化");

                // ========== 8. 保存输出 ==========
                string fileName = Path.Combine(AppContext.BaseDirectory, $@"SalesReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                workbook.SaveAs(fileName);

                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("  ✓ 报表生成成功！");
                Console.WriteLine("========================================");
                Console.WriteLine();
                Console.WriteLine($"文件已保存到: {fileName}");
                Console.WriteLine();
                Console.WriteLine("生成的报表包含以下工作表:");
                Console.WriteLine("  1. 源数据 - 原始销售数据（32条记录）");
                Console.WriteLine("  2. 产品销售分析 - 按产品和地区统计销售情况");
                Console.WriteLine("  3. 地区销售分析 - 各地区销售对比");
                Console.WriteLine("  4. 销售人员业绩 - 销售人员业绩排名");
                Console.WriteLine("  5. 月度销售趋势 - 月度销售变化趋势");
                Console.WriteLine();
                Console.WriteLine("数据分析要点:");
                Console.WriteLine("  - 涵盖4个销售地区：北京、上海、广州、深圳");
                Console.WriteLine("  - 包含3个产品类别：电子产品、家居用品、服装");
                Console.WriteLine("  - 记录4名销售人员业绩：张三、李四、王五、赵六");
                Console.WriteLine("  - 跨度时间：2023年1月至12月");
                Console.WriteLine();
                Console.WriteLine("按任意键退出程序...");
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("  ✗ 生成报表时出错！");
                Console.WriteLine("========================================");
                Console.WriteLine();
                Console.WriteLine($"错误信息: {ex.Message}");
                Console.WriteLine();
                Console.WriteLine($"详细信息: {ex.StackTrace}");
                Console.WriteLine();
                Console.WriteLine("按任意键退出...");
            }

            Console.ReadKey();
        }

        /// <summary>
        /// 添加报表标题
        /// </summary>
        /// <param name="worksheet">工作表对象</param>
        /// <param name="title">标题文本</param>
        static void AddReportTitle(IExcelWorksheet worksheet, string title)
        {
            worksheet.Rows[1].Insert();
            worksheet.Range("A1").Value = title;
            worksheet.Range("A1").Font.Size = 16;
            worksheet.Range("A1").Font.Bold = true;
            worksheet.Range("A1").Font.Color = Color.DarkBlue; // 深蓝色
            worksheet.Range("A1").ColumnWidth = 25;
        }
    }
}
