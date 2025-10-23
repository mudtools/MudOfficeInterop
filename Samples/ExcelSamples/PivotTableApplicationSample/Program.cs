//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;

namespace PivotTableApplicationSample
{
    /// <summary>
    /// 数据透视表应用示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel创建和配置数据透视表
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("数据透视表应用示例");
            Console.WriteLine("==================");
            Console.WriteLine();

            // 演示基础数据透视表创建
            BasicPivotTableExample();

            // 演示销售数据透视表创建
            SalesPivotTableExample();

            // 演示多维数据透视表创建
            MultiDimensionalPivotTableExample();

            // 演示带有计算字段的数据透视表创建
            CalculatedFieldPivotTableExample();

            // 演示数据透视表格式设置
            PivotTableFormattingExample();

            // 演示数据透视表筛选器应用
            PivotTableFilterExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 基础数据透视表创建示例
        /// 演示如何创建和配置基础数据透视表
        /// </summary>
        static void BasicPivotTableExample()
        {
            Console.WriteLine("=== 基础数据透视表创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var sourceWorksheet = workbook.ActiveSheetWrap;
                sourceWorksheet.Name = "源数据";

                // 创建销售数据
                sourceWorksheet.Range("A1").Value = "产品类别";
                sourceWorksheet.Range("B1").Value = "产品名称";
                sourceWorksheet.Range("C1").Value = "销售地区";
                sourceWorksheet.Range("D1").Value = "销售数量";
                sourceWorksheet.Range("E1").Value = "销售金额";

                object[,] salesData = {
                    {"电子产品", "笔记本电脑", "北京", 10, 50000},
                    {"电子产品", "台式电脑", "上海", 8, 32000},
                    {"电子产品", "平板电脑", "广州", 15, 30000},
                    {"电子产品", "手机", "深圳", 20, 60000},
                    {"家居用品", "沙发", "北京", 5, 15000},
                    {"家居用品", "餐桌", "上海", 3, 9000},
                    {"家居用品", "床", "广州", 4, 12000},
                    {"家居用品", "衣柜", "深圳", 6, 18000},
                    {"服装", "T恤", "北京", 50, 2500},
                    {"服装", "牛仔裤", "上海", 40, 4000},
                    {"服装", "外套", "广州", 30, 6000},
                    {"服装", "连衣裙", "深圳", 35, 7000}
                };

                var dataRange = sourceWorksheet.Range("A2:E13");
                dataRange.Value = salesData;

                // 创建数据透视表工作表
                var pivotWorksheet = workbook.Worksheets.Add() as IExcelWorksheet;
                pivotWorksheet.Name = "基础透视表";

                // 创建数据透视表缓存
                var pivotCache = pivotWorksheet.PivotCaches().Create(XlPivotTableSourceType.xlConsolidation, sourceWorksheet.Range("A1:E13"));

                // 创建数据透视表
                var pivotTable = pivotWorksheet.PivotTables().Add(pivotCache, pivotWorksheet.Range("A1"), "BasicPivotTable");

                // 配置字段
                // 添加行字段 - 产品类别
                pivotTable.PivotFields("产品类别").Orientation = XlPivotFieldOrientation.xlRowField;

                // 添加列字段 - 销售地区
                pivotTable.PivotFields("销售地区").Orientation = XlPivotFieldOrientation.xlColumnField;

                // 添加值字段 - 销售金额
                var dataField = pivotTable.PivotFields("销售金额");
                dataField.Orientation = XlPivotFieldOrientation.xlDataField;
                dataField.Function = XlConsolidationFunction.xlSum;

                // 自动调整列宽
                sourceWorksheet.Columns.AutoFit();
                pivotWorksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"BasicPivotTable_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建基础数据透视表: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建基础数据透视表时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 销售数据透视表创建示例
        /// 演示如何创建复杂的销售数据透视表
        /// </summary>
        static void SalesPivotTableExample()
        {
            Console.WriteLine("=== 销售数据透视表创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var sourceWorksheet = workbook.ActiveSheetWrap;
                sourceWorksheet.Name = "销售数据";

                // 创建详细销售数据
                sourceWorksheet.Range("A1").Value = "日期";
                sourceWorksheet.Range("B1").Value = "产品类别";
                sourceWorksheet.Range("C1").Value = "产品名称";
                sourceWorksheet.Range("D1").Value = "销售地区";
                sourceWorksheet.Range("E1").Value = "销售人员";
                sourceWorksheet.Range("F1").Value = "销售数量";
                sourceWorksheet.Range("G1").Value = "单价";
                sourceWorksheet.Range("H1").Value = "销售金额";

                object[,] detailedSalesData = {
                    {"2023-01-01", "电子产品", "笔记本电脑", "北京", "张三", 2, 5000, 10000},
                    {"2023-01-02", "电子产品", "台式电脑", "上海", "李四", 1, 4000, 4000},
                    {"2023-01-03", "家居用品", "沙发", "广州", "王五", 1, 3000, 3000},
                    {"2023-01-04", "服装", "T恤", "深圳", "赵六", 10, 50, 500},
                    {"2023-01-05", "电子产品", "手机", "北京", "张三", 3, 2000, 6000},
                    {"2023-01-06", "家居用品", "床", "上海", "李四", 1, 2000, 2000},
                    {"2023-01-07", "服装", "牛仔裤", "广州", "王五", 5, 100, 500},
                    {"2023-01-08", "电子产品", "平板电脑", "深圳", "赵六", 2, 1500, 3000},
                    {"2023-01-09", "家居用品", "餐桌", "北京", "张三", 1, 1000, 1000},
                    {"2023-01-10", "服装", "外套", "上海", "李四", 3, 300, 900},
                    {"2023-02-01", "电子产品", "笔记本电脑", "广州", "王五", 1, 5000, 5000},
                    {"2023-02-02", "电子产品", "手机", "深圳", "赵六", 2, 2000, 4000},
                    {"2023-02-03", "家居用品", "沙发", "北京", "张三", 2, 3000, 6000},
                    {"2023-02-04", "服装", "连衣裙", "上海", "李四", 4, 200, 800},
                    {"2023-02-05", "电子产品", "台式电脑", "广州", "王五", 2, 4000, 8000}
                };

                var dataRange = sourceWorksheet.Range("A2:H16");
                dataRange.Value = detailedSalesData;

                // 创建数据透视表工作表
                var pivotWorksheet = workbook.Worksheets.Add() as IExcelWorksheet;
                pivotWorksheet.Name = "销售透视表";

                // 创建数据透视表缓存
                var pivotCache = pivotWorksheet.PivotCaches().Create(XlPivotTableSourceType.xlConsolidation, sourceWorksheet.Range("A1:H16"));

                // 创建数据透视表
                var pivotTable = pivotWorksheet.PivotTables().Add(pivotCache, pivotWorksheet.Range("A1"), "SalesPivotTable");

                // 配置字段
                // 添加页字段 - 日期（作为筛选器）
                pivotTable.PivotFields("日期").Orientation = XlPivotFieldOrientation.xlPageField;

                // 添加行字段 - 产品类别
                pivotTable.PivotFields("产品类别").Orientation = XlPivotFieldOrientation.xlRowField;

                // 添加行字段 - 产品名称
                pivotTable.PivotFields("产品名称").Orientation = XlPivotFieldOrientation.xlRowField;

                // 添加列字段 - 销售地区
                pivotTable.PivotFields("销售地区").Orientation = XlPivotFieldOrientation.xlColumnField;

                // 添加值字段 - 销售金额（求和）
                var sumField = pivotTable.PivotFields("销售金额");
                sumField.Orientation = XlPivotFieldOrientation.xlDataField;
                sumField.Function = XlConsolidationFunction.xlSum;
                sumField.Name = "销售金额合计";

                // 添加值字段 - 销售数量（计数）
                var countField = pivotTable.PivotFields("销售数量");
                countField.Orientation = XlPivotFieldOrientation.xlDataField;
                countField.Function = XlConsolidationFunction.xlCount;
                countField.Name = "销售次数";

                // 设置数据透视表选项
                pivotTable.RowGrand = true;  // 显示行总计
                pivotTable.ColumnGrand = true;  // 显示列总计
                pivotTable.HasAutoFormat = true;  // 自动套用格式

                // 自动调整列宽
                sourceWorksheet.Columns.AutoFit();
                pivotWorksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"SalesPivotTable_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建销售数据透视表: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建销售数据透视表时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 多维数据透视表创建示例
        /// 演示如何创建多维数据透视表
        /// </summary>
        static void MultiDimensionalPivotTableExample()
        {
            Console.WriteLine("=== 多维数据透视表创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var sourceWorksheet = workbook.ActiveSheetWrap;
                sourceWorksheet.Name = "多维数据";

                // 创建季度销售数据
                sourceWorksheet.Range("A1").Value = "年份";
                sourceWorksheet.Range("B1").Value = "季度";
                sourceWorksheet.Range("C1").Value = "产品类别";
                sourceWorksheet.Range("D1").Value = "产品名称";
                sourceWorksheet.Range("E1").Value = "销售地区";
                sourceWorksheet.Range("F1").Value = "销售金额";

                object[,] quarterlyData = {
                    {"2022", "Q1", "电子产品", "笔记本电脑", "北京", 50000},
                    {"2022", "Q1", "电子产品", "手机", "北京", 30000},
                    {"2022", "Q1", "家居用品", "沙发", "北京", 15000},
                    {"2022", "Q2", "电子产品", "笔记本电脑", "上海", 60000},
                    {"2022", "Q2", "电子产品", "手机", "上海", 35000},
                    {"2022", "Q2", "家居用品", "沙发", "上海", 18000},
                    {"2022", "Q3", "电子产品", "笔记本电脑", "广州", 55000},
                    {"2022", "Q3", "电子产品", "手机", "广州", 40000},
                    {"2022", "Q3", "家居用品", "沙发", "广州", 20000},
                    {"2022", "Q4", "电子产品", "笔记本电脑", "深圳", 70000},
                    {"2022", "Q4", "电子产品", "手机", "深圳", 45000},
                    {"2022", "Q4", "家居用品", "沙发", "深圳", 25000},
                    {"2023", "Q1", "电子产品", "笔记本电脑", "北京", 75000},
                    {"2023", "Q1", "电子产品", "手机", "北京", 50000},
                    {"2023", "Q1", "家居用品", "沙发", "北京", 30000}
                };

                var dataRange = sourceWorksheet.Range("A2:F16");
                dataRange.Value = quarterlyData;

                // 创建数据透视表工作表
                var pivotWorksheet = workbook.Worksheets.Add() as IExcelWorksheet;
                pivotWorksheet.Name = "多维透视表";

                // 创建数据透视表缓存
                var pivotCache = pivotWorksheet.PivotCaches().Create(XlPivotTableSourceType.xlConsolidation, sourceWorksheet.Range("A1:F16"));

                // 创建数据透视表
                var pivotTable = pivotWorksheet.PivotTables().Add(pivotCache, pivotWorksheet.Range("A1"), "MultiDimensionalPivotTable");

                // 配置字段
                // 添加页字段 - 年份
                pivotTable.PivotFields("年份").Orientation = XlPivotFieldOrientation.xlPageField;

                // 添加行字段 - 季度
                pivotTable.PivotFields("季度").Orientation = XlPivotFieldOrientation.xlRowField;

                // 添加行字段 - 产品类别
                pivotTable.PivotFields("产品类别").Orientation = XlPivotFieldOrientation.xlRowField;

                // 添加列字段 - 销售地区
                pivotTable.PivotFields("销售地区").Orientation = XlPivotFieldOrientation.xlColumnField;

                // 添加值字段 - 销售金额
                var dataField = pivotTable.PivotFields("销售金额");
                dataField.Orientation = XlPivotFieldOrientation.xlDataField;
                dataField.Function = XlConsolidationFunction.xlSum;
                dataField.Name = "销售金额合计";

                // 设置数据透视表选项
                pivotTable.RowGrand = true;
                pivotTable.ColumnGrand = true;
                pivotTable.HasAutoFormat = true;

                // 自动调整列宽
                sourceWorksheet.Columns.AutoFit();
                pivotWorksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"MultiDimensionalPivotTable_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建多维数据透视表: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建多维数据透视表时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 带有计算字段的数据透视表创建示例
        /// 演示如何在数据透视表中添加计算字段
        /// </summary>
        static void CalculatedFieldPivotTableExample()
        {
            Console.WriteLine("=== 带有计算字段的数据透视表创建示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var sourceWorksheet = workbook.ActiveSheetWrap;
                sourceWorksheet.Name = "计算字段数据";

                // 创建产品成本数据
                sourceWorksheet.Range("A1").Value = "产品类别";
                sourceWorksheet.Range("B1").Value = "产品名称";
                sourceWorksheet.Range("C1").Value = "销售金额";
                sourceWorksheet.Range("D1").Value = "成本金额";

                object[,] productData = {
                    {"电子产品", "笔记本电脑", 50000, 40000},
                    {"电子产品", "手机", 30000, 24000},
                    {"电子产品", "平板电脑", 20000, 16000},
                    {"家居用品", "沙发", 15000, 10000},
                    {"家居用品", "床", 12000, 8000},
                    {"家居用品", "餐桌", 10000, 6000},
                    {"服装", "T恤", 5000, 2000},
                    {"服装", "牛仔裤", 8000, 3200},
                    {"服装", "外套", 12000, 4800}
                };

                var dataRange = sourceWorksheet.Range("A2:D10");
                dataRange.Value = productData;

                // 创建数据透视表工作表
                var pivotWorksheet = workbook.Worksheets.Add() as IExcelWorksheet;
                pivotWorksheet.Name = "计算字段透视表";

                // 创建数据透视表缓存
                var pivotCache = pivotWorksheet.PivotCaches().Create(XlPivotTableSourceType.xlConsolidation, sourceWorksheet.Range("A1:D10"));

                // 创建数据透视表
                var pivotTable = pivotWorksheet.PivotTables().Add(pivotCache, pivotWorksheet.Range("A1"), "CalculatedFieldPivotTable");

                // 配置字段
                // 添加行字段 - 产品类别
                pivotTable.PivotFields("产品类别").Orientation = XlPivotFieldOrientation.xlRowField;

                // 添加值字段 - 销售金额
                var salesField = pivotTable.PivotFields("销售金额");
                salesField.Orientation = XlPivotFieldOrientation.xlDataField;
                salesField.Function = XlConsolidationFunction.xlSum;
                salesField.Name = "销售金额";

                // 添加值字段 - 成本金额
                var costField = pivotTable.PivotFields("成本金额");
                costField.Orientation = XlPivotFieldOrientation.xlDataField;
                costField.Function = XlConsolidationFunction.xlSum;
                costField.Name = "成本金额";

                // 添加计算字段 - 利润
                pivotTable.CalculatedFields().Add("利润", "=销售金额-成本金额");

                // 添加计算字段 - 利润率
                pivotTable.CalculatedFields().Add("利润率", "=利润/销售金额");

                // 设置数据透视表选项
                pivotTable.RowGrand = true;
                pivotTable.ColumnGrand = true;

                // 格式化利润率字段为百分比
                var profitMarginField = pivotTable.PivotFields("利润率");
                // 注意：这里需要根据实际API调整格式设置方式

                // 自动调整列宽
                sourceWorksheet.Columns.AutoFit();
                pivotWorksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"CalculatedFieldPivotTable_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建带有计算字段的数据透视表: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建带有计算字段的数据透视表时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据透视表格式设置示例
        /// 演示如何设置数据透视表的格式
        /// </summary>
        static void PivotTableFormattingExample()
        {
            Console.WriteLine("=== 数据透视表格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var sourceWorksheet = workbook.ActiveSheetWrap;
                sourceWorksheet.Name = "格式数据";

                // 创建销售数据
                sourceWorksheet.Range("A1").Value = "产品类别";
                sourceWorksheet.Range("B1").Value = "产品名称";
                sourceWorksheet.Range("C1").Value = "销售地区";
                sourceWorksheet.Range("D1").Value = "销售金额";

                object[,] salesData = {
                    {"电子产品", "笔记本电脑", "北京", 50000},
                    {"电子产品", "手机", "上海", 30000},
                    {"电子产品", "平板电脑", "广州", 20000},
                    {"家居用品", "沙发", "深圳", 15000},
                    {"家居用品", "床", "北京", 12000},
                    {"家居用品", "餐桌", "上海", 10000},
                    {"服装", "T恤", "广州", 5000},
                    {"服装", "牛仔裤", "深圳", 8000},
                    {"服装", "外套", "北京", 12000}
                };

                var dataRange = sourceWorksheet.Range("A2:D10");
                dataRange.Value = salesData;

                // 创建数据透视表工作表
                var pivotWorksheet = workbook.Worksheets.Add() as IExcelWorksheet;
                pivotWorksheet.Name = "格式透视表";

                // 创建数据透视表缓存
                var pivotCache = pivotWorksheet.PivotCaches().Create(XlPivotTableSourceType.xlConsolidation, sourceWorksheet.Range("A1:D10"));

                // 创建数据透视表
                var pivotTable = pivotWorksheet.PivotTables().Add(pivotCache, pivotWorksheet.Range("A1"), "FormattingPivotTable");

                // 配置字段
                // 添加行字段 - 产品类别
                pivotTable.PivotFields("产品类别").Orientation = XlPivotFieldOrientation.xlRowField;

                // 添加行字段 - 产品名称
                pivotTable.PivotFields("产品名称").Orientation = XlPivotFieldOrientation.xlRowField;

                // 添加列字段 - 销售地区
                pivotTable.PivotFields("销售地区").Orientation = XlPivotFieldOrientation.xlColumnField;

                // 添加值字段 - 销售金额
                var dataField = pivotTable.PivotFields("销售金额");
                dataField.Orientation = XlPivotFieldOrientation.xlDataField;
                dataField.Function = XlConsolidationFunction.xlSum;
                dataField.NumberFormat = "¥#,##0";

                // 设置数据透视表选项
                pivotTable.ShowTableStyleRowStripes = true;
                pivotTable.ShowTableStyleColumnStripes = true;
                pivotTable.ShowTableStyleLastColumn = true;

                // 格式化总计行
                pivotTable.RowGrand = true;
                pivotTable.ColumnGrand = true;

                // 自动调整列宽
                sourceWorksheet.Columns.AutoFit();
                pivotWorksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"FormattingPivotTable_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建格式化数据透视表: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建格式化数据透视表时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据透视表筛选器应用示例
        /// 演示如何使用数据透视表的筛选功能
        /// </summary>
        static void PivotTableFilterExample()
        {
            Console.WriteLine("=== 数据透视表筛选器应用示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var sourceWorksheet = workbook.ActiveSheetWrap;
                sourceWorksheet.Name = "筛选数据";

                // 创建详细销售数据
                sourceWorksheet.Range("A1").Value = "日期";
                sourceWorksheet.Range("B1").Value = "产品类别";
                sourceWorksheet.Range("C1").Value = "产品名称";
                sourceWorksheet.Range("D1").Value = "销售地区";
                sourceWorksheet.Range("E1").Value = "销售人员";
                sourceWorksheet.Range("F1").Value = "销售金额";

                object[,] detailedSalesData = {
                    {"2023-01-01", "电子产品", "笔记本电脑", "北京", "张三", 10000},
                    {"2023-01-02", "电子产品", "手机", "上海", "李四", 6000},
                    {"2023-01-03", "家居用品", "沙发", "广州", "王五", 3000},
                    {"2023-01-04", "服装", "T恤", "深圳", "赵六", 500},
                    {"2023-01-05", "电子产品", "平板电脑", "北京", "张三", 3000},
                    {"2023-02-01", "电子产品", "笔记本电脑", "上海", "李四", 15000},
                    {"2023-02-02", "家居用品", "床", "广州", "王五", 2000},
                    {"2023-02-03", "服装", "牛仔裤", "深圳", "赵六", 800},
                    {"2023-02-04", "电子产品", "手机", "北京", "张三", 9000},
                    {"2023-03-01", "电子产品", "笔记本电脑", "广州", "王五", 20000},
                    {"2023-03-02", "家居用品", "餐桌", "深圳", "赵六", 1000},
                    {"2023-03-03", "服装", "外套", "北京", "张三", 1200}
                };

                var dataRange = sourceWorksheet.Range("A2:F13");
                dataRange.Value = detailedSalesData;

                // 创建数据透视表工作表
                var pivotWorksheet = workbook.Worksheets.Add() as IExcelWorksheet;
                pivotWorksheet.Name = "筛选透视表";

                // 创建数据透视表缓存
                var pivotCache = pivotWorksheet.PivotCaches().Create(XlPivotTableSourceType.xlConsolidation, sourceWorksheet.Range("A1:F13"));

                // 创建数据透视表
                var pivotTable = pivotWorksheet.PivotTables().Add(pivotCache, pivotWorksheet.Range("A1"), "FilterPivotTable");

                // 配置字段
                // 添加页字段 - 销售人员（作为筛选器）
                pivotTable.PivotFields("销售人员").Orientation = XlPivotFieldOrientation.xlPageField;

                // 添加页字段 - 产品类别（作为筛选器）
                pivotTable.PivotFields("产品类别").Orientation = XlPivotFieldOrientation.xlPageField;

                // 添加行字段 - 日期
                pivotTable.PivotFields("日期").Orientation = XlPivotFieldOrientation.xlRowField;

                // 添加列字段 - 销售地区
                pivotTable.PivotFields("销售地区").Orientation = XlPivotFieldOrientation.xlColumnField;

                // 添加值字段 - 销售金额
                var dataField = pivotTable.PivotFields("销售金额");
                dataField.Orientation = XlPivotFieldOrientation.xlDataField;
                dataField.Function = XlConsolidationFunction.xlSum;
                dataField.NumberFormat = "¥#,##0";

                // 自动调整列宽
                sourceWorksheet.Columns.AutoFit();
                pivotWorksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"FilterPivotTable_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建筛选数据透视表: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建筛选数据透视表时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}