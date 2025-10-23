//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;
using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;

namespace AdvancedDataOperationsSample
{
    /// <summary>
    /// 高级数据操作示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行高级数据操作
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("高级数据操作示例");
            Console.WriteLine("==============");
            Console.WriteLine();

            // 演示数据排序技术
            DataSortingExample();

            // 演示数据筛选功能
            DataFilteringExample();

            // 演示数据分组功能
            DataGroupingExample();

            // 演示分类汇总功能
            SubtotalExample();

            // 演示高级筛选功能
            AdvancedFilterExample();

            // 演示数据去重功能
            RemoveDuplicatesExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 数据排序技术示例
        /// 演示如何对数据进行单列和多列排序
        /// </summary>
        static void DataSortingExample()
        {
            Console.WriteLine("=== 数据排序技术示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数据排序";

                // 创建销售数据
                worksheet.Range("A1").Value = "产品类别";
                worksheet.Range("B1").Value = "产品名称";
                worksheet.Range("C1").Value = "销售地区";
                worksheet.Range("D1").Value = "销售数量";
                worksheet.Range("E1").Value = "销售金额";

                object[,] salesData = {
                    {"电子产品", "笔记本电脑", "北京", 10, 50000},
                    {"家居用品", "沙发", "上海", 5, 15000},
                    {"服装", "T恤", "广州", 50, 2500},
                    {"电子产品", "手机", "深圳", 20, 60000},
                    {"家居用品", "床", "北京", 3, 9000},
                    {"服装", "牛仔裤", "上海", 30, 3000},
                    {"电子产品", "平板电脑", "广州", 15, 30000},
                    {"家居用品", "餐桌", "深圳", 4, 12000},
                    {"服装", "外套", "北京", 20, 4000},
                    {"电子产品", "台式电脑", "上海", 8, 32000}
                };

                var dataRange = worksheet.Range("A2:E11");
                dataRange.Value = salesData;

                // 单列排序 - 按销售金额降序排序
                var sort = worksheet.Sort;
                sort.SetRange(dataRange);
                sort.Header = XlYesNoGuess.xlYes;
                sort.Orientation = XlSortOrientation.xlSortColumns;
                sort.SortFields.Clear();
                sort.SortFields.Add(worksheet.Range("E2:E11"), XlSortOn.xlSortOnValues, XlSortOrder.xlDescending);
                sort.Apply();

                worksheet.Range("G1").Value = "单列排序结果（按销售金额降序）";
                worksheet.Range("G1").Font.Bold = true;
                worksheet.Range("G1").Interior.Color = Color.LightBlue;

                // 复制排序后的数据到G列
                worksheet.Range("A1:E11").Copy(worksheet.Range("G1"));

                // 多列排序 - 先按产品类别升序，再按销售金额降序
                // 重新填充原始数据
                dataRange.Value = salesData;

                sort.SetRange(dataRange);
                sort.Header = XlYesNoGuess.xlYes;
                sort.Orientation = XlSortOrientation.xlSortColumns;
                sort.SortFields.Clear();
                sort.SortFields.Add(worksheet.Range("A2:A11"), XlSortOn.xlSortOnValues, XlSortOrder.xlAscending);
                sort.SortFields.Add(worksheet.Range("E2:E11"), XlSortOn.xlSortOnValues, XlSortOrder.xlDescending);
                sort.Apply();

                worksheet.Range("L1").Value = "多列排序结果（产品类别升序，销售金额降序）";
                worksheet.Range("L1").Font.Bold = true;
                worksheet.Range("L1").Interior.Color = Color.LightGreen;

                // 复制排序后的数据到L列
                worksheet.Range("A1:E11").Copy(worksheet.Range("L1"));

                // 设置数字格式
                worksheet.Range("D2:D11").NumberFormat = "0";
                worksheet.Range("E2:E11").NumberFormat = "¥#,##0";
                worksheet.Range("K2:K11").NumberFormat = "0";
                worksheet.Range("L2:L11").NumberFormat = "¥#,##0";
                worksheet.Range("P2:P11").NumberFormat = "0";
                worksheet.Range("Q2:Q11").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"DataSorting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数据排序技术: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数据排序技术时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据筛选功能示例
        /// 演示如何对数据进行自动筛选
        /// </summary>
        static void DataFilteringExample()
        {
            Console.WriteLine("=== 数据筛选功能示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数据筛选";

                // 创建销售数据
                worksheet.Range("A1").Value = "产品类别";
                worksheet.Range("B1").Value = "产品名称";
                worksheet.Range("C1").Value = "销售地区";
                worksheet.Range("D1").Value = "销售数量";
                worksheet.Range("E1").Value = "销售金额";

                object[,] salesData = {
                    {"电子产品", "笔记本电脑", "北京", 10, 50000},
                    {"家居用品", "沙发", "上海", 5, 15000},
                    {"服装", "T恤", "广州", 50, 2500},
                    {"电子产品", "手机", "深圳", 20, 60000},
                    {"家居用品", "床", "北京", 3, 9000},
                    {"服装", "牛仔裤", "上海", 30, 3000},
                    {"电子产品", "平板电脑", "广州", 15, 30000},
                    {"家居用品", "餐桌", "深圳", 4, 12000},
                    {"服装", "外套", "北京", 20, 4000},
                    {"电子产品", "台式电脑", "上海", 8, 32000}
                };

                var dataRange = worksheet.Range("A2:E11");
                dataRange.Value = salesData;

                // 添加自动筛选
                dataRange.AutoFilter();

                // 筛选电子产品
                worksheet.Range("A1").AutoFilter(1, "电子产品");

                // 复制筛选结果到另一位置
                worksheet.Range("G1").Value = "电子产品筛选结果";
                worksheet.Range("G1").Font.Bold = true;
                worksheet.Range("G1").Interior.Color = Color.LightBlue;

                // 由于筛选是动态的，我们创建一个新工作表来展示筛选结果
                var filteredSheet = workbook.Worksheets.Add() as IExcelWorksheet;
                filteredSheet.Name = "筛选结果";

                // 复制标题行
                worksheet.Range("A1:E1").Copy(filteredSheet.Range("A1"));

                // 手动复制符合条件的数据（模拟筛选结果）
                object[,] electronicData = {
                    {"电子产品", "笔记本电脑", "北京", 10, 50000},
                    {"电子产品", "手机", "深圳", 20, 60000},
                    {"电子产品", "平板电脑", "广州", 15, 30000},
                    {"电子产品", "台式电脑", "上海", 8, 32000}
                };

                filteredSheet.Range("A2:E5").Value = electronicData;

                // 设置数字格式
                worksheet.Range("D2:D11").NumberFormat = "0";
                worksheet.Range("E2:E11").NumberFormat = "¥#,##0";
                filteredSheet.Range("D2:D5").NumberFormat = "0";
                filteredSheet.Range("E2:E5").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();
                filteredSheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"DataFiltering_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数据筛选功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数据筛选功能时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据分组功能示例
        /// 演示如何对数据进行分组操作
        /// </summary>
        static void DataGroupingExample()
        {
            Console.WriteLine("=== 数据分组功能示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数据分组";

                // 创建详细销售数据
                worksheet.Range("A1").Value = "日期";
                worksheet.Range("B1").Value = "产品类别";
                worksheet.Range("C1").Value = "产品名称";
                worksheet.Range("D1").Value = "销售数量";
                worksheet.Range("E1").Value = "销售金额";

                object[,] salesData = {
                    {"2023-01-01", "电子产品", "笔记本电脑", 2, 10000},
                    {"2023-01-01", "电子产品", "手机", 5, 15000},
                    {"2023-01-02", "家居用品", "沙发", 1, 3000},
                    {"2023-01-02", "服装", "T恤", 10, 500},
                    {"2023-01-03", "电子产品", "平板电脑", 3, 6000},
                    {"2023-01-03", "家居用品", "床", 1, 3000},
                    {"2023-01-04", "服装", "牛仔裤", 5, 500},
                    {"2023-01-04", "电子产品", "手机", 3, 9000},
                    {"2023-01-05", "家居用品", "餐桌", 1, 3000},
                    {"2023-01-05", "服装", "外套", 2, 400}
                };

                var dataRange = worksheet.Range("A2:E11");
                dataRange.Value = salesData;

                // 按日期分组
                worksheet.Rows["3:3"].Group(); // 1月2日数据
                worksheet.Rows["5:5"].Group(); // 1月3日数据
                worksheet.Rows["7:7"].Group(); // 1月4日数据
                worksheet.Rows["9:9"].Group(); // 1月5日数据

                // 设置分组级别
                worksheet.Outline.AutomaticStyles = false;
                worksheet.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;
                worksheet.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnLeft;

                // 展开所有分组
                worksheet.Outline.ShowLevels(1, 1);

                // 设置数字格式
                worksheet.Range("D2:D11").NumberFormat = "0";
                worksheet.Range("E2:E11").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"DataGrouping_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数据分组功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数据分组功能时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 分类汇总功能示例
        /// 演示如何对数据进行分类汇总操作
        /// </summary>
        static void SubtotalExample()
        {
            Console.WriteLine("=== 分类汇总功能示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "分类汇总";

                // 创建销售数据（已按产品类别排序）
                worksheet.Range("A1").Value = "产品类别";
                worksheet.Range("B1").Value = "产品名称";
                worksheet.Range("C1").Value = "销售数量";
                worksheet.Range("D1").Value = "销售金额";

                object[,] salesData = {
                    {"电子产品", "笔记本电脑", 10, 50000},
                    {"电子产品", "手机", 20, 60000},
                    {"电子产品", "平板电脑", 15, 30000},
                    {"电子产品", "台式电脑", 8, 32000},
                    {"家居用品", "沙发", 5, 15000},
                    {"家居用品", "床", 3, 9000},
                    {"家居用品", "餐桌", 4, 12000},
                    {"服装", "T恤", 50, 2500},
                    {"服装", "牛仔裤", 30, 3000},
                    {"服装", "外套", 20, 4000}
                };

                var dataRange = worksheet.Range("A2:D11");
                dataRange.Value = salesData;

                // 添加分类汇总
                // 首先按产品类别排序（分类汇总要求数据已排序）
                var sort = worksheet.Sort;
                sort.SetRange(dataRange);
                sort.Header = XlYesNoGuess.xlYes;
                sort.Orientation = XlSortOrientation.xlSortColumns;
                sort.SortFields.Clear();
                sort.SortFields.Add(worksheet.Range("A2:A11"), XlSortOn.xlSortOnValues, XlSortOrder.xlAscending);
                sort.Apply();

                // 应用分类汇总
                dataRange.Subtotal(1, XlConsolidationFunction.xlSum, new int[] { 3, 4 }, true, false, XlSummaryRow.xlSummaryBelow);

                // 设置数字格式
                worksheet.Range("C2:D15").NumberFormat = "0";
                worksheet.Range("D2:D15").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"Subtotal_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示分类汇总功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 分类汇总功能时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 高级筛选功能示例
        /// 演示如何使用高级筛选功能
        /// </summary>
        static void AdvancedFilterExample()
        {
            Console.WriteLine("=== 高级筛选功能示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "高级筛选";

                // 创建销售数据
                worksheet.Range("A1").Value = "产品类别";
                worksheet.Range("B1").Value = "产品名称";
                worksheet.Range("C1").Value = "销售地区";
                worksheet.Range("D1").Value = "销售金额";

                object[,] salesData = {
                    {"电子产品", "笔记本电脑", "北京", 50000},
                    {"家居用品", "沙发", "上海", 15000},
                    {"服装", "T恤", "广州", 2500},
                    {"电子产品", "手机", "深圳", 60000},
                    {"家居用品", "床", "北京", 9000},
                    {"服装", "牛仔裤", "上海", 3000},
                    {"电子产品", "平板电脑", "广州", 30000},
                    {"家居用品", "餐桌", "深圳", 12000},
                    {"服装", "外套", "北京", 4000},
                    {"电子产品", "台式电脑", "上海", 32000}
                };

                var dataRange = worksheet.Range("A2:D11");
                dataRange.Value = salesData;

                // 设置筛选条件区域
                worksheet.Range("F1").Value = "产品类别";
                worksheet.Range("F2").Value = "电子产品";
                worksheet.Range("F3").Value = "家居用品";

                worksheet.Range("G1").Value = "销售金额";
                worksheet.Range("G2").Value = ">30000";

                // 执行高级筛选 - 筛选结果复制到新位置
                var criteriaRange = worksheet.Range("F1:G3");
                dataRange.AdvancedFilter(XlFilterAction.xlFilterCopy, criteriaRange, worksheet.Range("I1"));

                worksheet.Range("I1").Value = "高级筛选结果";
                worksheet.Range("I1").Font.Bold = true;
                worksheet.Range("I1").Interior.Color = Color.LightBlue;

                // 设置数字格式
                worksheet.Range("D2:D11").NumberFormat = "¥#,##0";
                worksheet.Range("I2:I11").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"AdvancedFilter_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示高级筛选功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 高级筛选功能时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据去重功能示例
        /// 演示如何对数据进行去重操作
        /// </summary>
        static void RemoveDuplicatesExample()
        {
            Console.WriteLine("=== 数据去重功能示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数据去重";

                // 创建包含重复数据的销售数据
                worksheet.Range("A1").Value = "产品类别";
                worksheet.Range("B1").Value = "产品名称";
                worksheet.Range("C1").Value = "销售地区";
                worksheet.Range("D1").Value = "销售金额";

                object[,] salesDataWithDuplicates = {
                    {"电子产品", "笔记本电脑", "北京", 50000},
                    {"家居用品", "沙发", "上海", 15000},
                    {"服装", "T恤", "广州", 2500},
                    {"电子产品", "手机", "深圳", 60000},
                    {"家居用品", "床", "北京", 9000},
                    {"服装", "牛仔裤", "上海", 3000},
                    {"电子产品", "笔记本电脑", "北京", 50000}, // 重复数据
                    {"家居用品", "沙发", "上海", 15000},       // 重复数据
                    {"服装", "外套", "北京", 4000},
                    {"电子产品", "台式电脑", "上海", 32000}
                };

                var dataRange = worksheet.Range("A2:D11");
                dataRange.Value = salesDataWithDuplicates;

                // 复制原始数据用于对比
                worksheet.Range("A1:D11").Copy(worksheet.Range("F1"));
                worksheet.Range("F1").Value = "原始数据（含重复）";
                worksheet.Range("F1").Font.Bold = true;
                worksheet.Range("F1").Interior.Color = Color.LightYellow;

                // 去除重复数据（基于所有列）
                dataRange.RemoveDuplicates(new int[] { 1, 2, 3, 4 }, XlYesNoGuess.xlYes);

                worksheet.Range("A1").Value = "去重后数据";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.LightGreen;

                // 设置数字格式
                worksheet.Range("D2:D11").NumberFormat = "¥#,##0";
                worksheet.Range("H2:H11").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"RemoveDuplicates_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数据去重功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数据去重功能时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}