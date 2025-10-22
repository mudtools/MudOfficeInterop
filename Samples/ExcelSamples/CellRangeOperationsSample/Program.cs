//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace CellRangeOperationsSample
{
    /// <summary>
    /// 单元格和区域操作示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行单元格和区域的基本操作
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("单元格和区域操作示例");
            Console.WriteLine("====================");
            Console.WriteLine();

            // 演示单元格基础操作
            CellBasicOperationsExample();

            // 演示区域基础操作
            RangeBasicOperationsExample();

            // 演示单元格数据读写操作
            CellDataOperationsExample();

            // 演示区域数据读写操作
            RangeDataOperationsExample();

            // 演示高级单元格和区域操作
            AdvancedCellRangeOperationsExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 单元格基础操作示例
        /// 演示如何引用和操作单个单元格
        /// </summary>
        static void CellBasicOperationsExample()
        {
            Console.WriteLine("=== 单元格基础操作示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "单元格操作";

                // 方法1：使用Range方法引用单元格
                var cellA1 = worksheet.Range("A1");
                cellA1.Value = "使用Range方法引用";
                cellA1.Font.Bold = true;
                cellA1.Interior.Color = Color.LightBlue;

                // 方法2：使用索引器（字符串地址）引用单元格
                var cellB1 = worksheet["B1"];
                cellB1.Value = "使用字符串索引器引用";
                cellB1.Font.Italic = true;
                cellB1.Interior.Color = Color.LightGreen;

                // 方法3：使用索引器（行列号）引用单元格
                var cellC1 = worksheet[1, 3]; // 第1行，第3列（C列）
                cellC1.Value = "使用行列号索引器引用";
                cellC1.Font.Underline = XlUnderlineStyle.xlUnderlineStyleSingle;
                cellC1.Interior.Color = Color.LightYellow;

                // 方法4：使用Cells属性引用单元格
                var cellD1 = worksheet.Cells[1, 4]; // 第1行，第4列（D列）
                cellD1.Value = "使用Cells属性引用";
                cellD1.Font.Color = Color.Red;
                cellD1.Interior.Color = Color.LightPink;

                // 相对引用示例
                worksheet["A3"].Value = "相对引用示例";
                var activeCell = worksheet["A3"];
                var offsetCell = activeCell.Offset(1, 1); // 向下1行，向右1列
                offsetCell.Value = "这是相对于A3的单元格";
                offsetCell.Interior.Color = Color.LightGray;

                // 保存工作簿
                string fileName = $"CellBasicOperations_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示单元格基础操作: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 单元格基础操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 区域基础操作示例
        /// 演示如何引用和操作单元格区域
        /// </summary>
        static void RangeBasicOperationsExample()
        {
            Console.WriteLine("=== 区域基础操作示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "区域操作";

                // 方法1：使用Range方法引用区域
                var rangeA1C3 = worksheet.Range("A1:C3");
                rangeA1C3.Value = "区域A1:C3";
                rangeA1C3.Font.Bold = true;
                rangeA1C3.Interior.Color = Color.LightBlue;

                // 方法2：使用索引器引用区域
                var rangeE1G3 = worksheet["E1:G3"];
                rangeE1G3.Value = "区域E1:G3";
                rangeE1G3.Font.Italic = true;
                rangeE1G3.Interior.Color = Color.LightGreen;

                // 方法3：使用Cells属性引用区域
                var rangeA5C7 = worksheet.Cells[5, 1, 7, 3]; // 第5-7行，第1-3列
                rangeA5C7.Value = "区域A5:C7";
                rangeA5C7.Font.Underline = true;
                rangeA5C7.Interior.Color = Color.LightYellow;

                // 使用Range方法定义区域范围
                var rangeFromTo = worksheet.Range(worksheet["A9"], worksheet["C11"]);
                rangeFromTo.Value = "从A9到C11的区域";
                rangeFromTo.Font.Color = Color.Red;
                rangeFromTo.Interior.Color = Color.LightPink;

                // 保存工作簿
                string fileName = $"RangeBasicOperations_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示区域基础操作: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 区域基础操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 单元格数据读写操作示例
        /// 演示如何读写不同类型的单元格数据
        /// </summary>
        static void CellDataOperationsExample()
        {
            Console.WriteLine("=== 单元格数据读写操作示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "单元格数据操作";

                // 写入不同类型的数据
                worksheet["A1"].Value = "文本数据";
                worksheet["A2"].Value = 123.45;           // 数字
                worksheet["A3"].Value = DateTime.Now;     // 日期时间
                worksheet["A4"].Value = true;             // 布尔值

                // 读取数据并显示
                Console.WriteLine($"A1单元格值: {worksheet["A1"].Value} (类型: {worksheet["A1"].Value?.GetType()})");
                Console.WriteLine($"A2单元格值: {worksheet["A2"].Value} (类型: {worksheet["A2"].Value?.GetType()})");
                Console.WriteLine($"A3单元格值: {worksheet["A3"].Value} (类型: {worksheet["A3"].Value?.GetType()})");
                Console.WriteLine($"A4单元格值: {worksheet["A4"].Value} (类型: {worksheet["A4"].Value?.GetType()})");

                // 设置公式
                worksheet["B1"].Value = 10;
                worksheet["B2"].Value = 20;
                worksheet["B3"].Formula = "=B1+B2";
                Console.WriteLine($"B3单元格公式结果: {worksheet["B3"].Value}");

                // 保存工作簿
                string fileName = $"CellDataOperations_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示单元格数据读写操作: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 单元格数据读写操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 区域数据读写操作示例
        /// 演示如何读写区域数据
        /// </summary>
        static void RangeDataOperationsExample()
        {
            Console.WriteLine("=== 区域数据读写操作示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                worksheet.Name = "区域数据操作";

                // 写入区域数据（使用二维数组）
                object[,] data = {
                    {"姓名", "部门", "薪资"},
                    {"张三", "技术部", 8000},
                    {"李四", "销售部", 7500},
                    {"王五", "市场部", 7000}
                };

                var dataRange = worksheet.Range("A1:C4");
                dataRange.Value = data;
                dataRange.Font.Bold = true;
                dataRange.Interior.Color = Color.LightBlue;

                // 读取区域数据
                object[,] readData = (object[,])dataRange.Value;
                Console.WriteLine("读取的区域数据:");
                for (int row = 0; row < readData.GetLength(0); row++)
                {
                    for (int col = 0; col < readData.GetLength(1); col++)
                    {
                        Console.Write($"{readData[row, col]}\t");
                    }
                    Console.WriteLine();
                }

                // 使用区域公式
                worksheet.Range("D1").Value = "奖金";
                worksheet.Range("D2:D4").Formula = "=C2*0.1";

                // 保存工作簿
                string fileName = $"RangeDataOperations_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示区域数据读写操作: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 区域数据读写操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 高级单元格和区域操作示例
        /// 演示更复杂的单元格和区域操作
        /// </summary>
        static void AdvancedCellRangeOperationsExample()
        {
            Console.WriteLine("=== 高级单元格和区域操作示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "高级操作";

                // 创建示例数据
                object[,] salesData = {
                    {"月份", "销售额", "利润"},
                    {"1月", 50000, 10000},
                    {"2月", 55000, 12000},
                    {"3月", 60000, 15000},
                    {"4月", 58000, 13000},
                    {"5月", 62000, 16000},
                    {"6月", 65000, 18000}
                };

                var dataRange = worksheet.Range("A1:C7");
                dataRange.Value = salesData;

                // 设置标题格式
                var headerRange = worksheet.Range("A1:C1");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 设置数据格式
                var numberRange = worksheet.Range("B2:C7");
                numberRange.NumberFormat = "#,##0";

                // 使用特殊区域引用
                var usedRange = worksheet.UsedRange;
                usedRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 查找操作
                var foundCell = dataRange.Find("6月");
                if (foundCell != null)
                {
                    foundCell.Interior.Color = Color.Yellow;
                    Console.WriteLine($"找到'6月'在单元格: {foundCell.Address}");
                }

                // 保存工作簿
                string fileName = $"AdvancedCellRangeOperations_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示高级单元格和区域操作: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 高级单元格和区域操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}