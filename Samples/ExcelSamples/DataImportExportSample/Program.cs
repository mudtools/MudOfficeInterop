//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace DataImportExportSample
{
    /// <summary>
    /// 数据导入导出与转换示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行数据导入导出操作
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("数据导入导出与转换示例");
            Console.WriteLine("======================");
            Console.WriteLine();

            // 演示从数组导入数据
            ImportFromArrayExample();

            // 演示导出数据到数组
            ExportToArrayExample();

            // 演示从CSV文件导入数据
            ImportFromCsvExample();

            // 演示导出数据到CSV文件
            ExportToCsvExample();

            // 演示数据格式转换
            DataFormatConversionExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 从数组导入数据示例
        /// 演示如何将内存中的数据导入到Excel工作表
        /// </summary>
        static void ImportFromArrayExample()
        {
            Console.WriteLine("=== 从数组导入数据示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数组导入数据";

                // 创建示例数据（二维数组）
                object[,] employeeData = {
                    {"员工ID", "姓名", "部门", "职位", "入职日期", "薪资"},
                    {1001, "张三", "技术部", "软件工程师", new DateTime(2020, 1, 15), 8000},
                    {1002, "李四", "销售部", "销售经理", new DateTime(2019, 3, 22), 12000},
                    {1003, "王五", "市场部", "市场专员", new DateTime(2021, 7, 10), 7000},
                    {1004, "赵六", "人事部", "人事专员", new DateTime(2020, 11, 5), 6500},
                    {1005, "钱七", "财务部", "会计师", new DateTime(2018, 9, 18), 9000}
                };

                // 将数据导入到工作表
                var dataRange = worksheet.Range("A1:F6");
                dataRange.Value = employeeData;

                // 设置标题格式
                var headerRange = worksheet.Range("A1:F1");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightBlue;
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 设置数据格式
                var dateRange = worksheet.Range("E2:E6");
                dateRange.NumberFormat = "yyyy-mm-dd";

                var salaryRange = worksheet.Range("F2:F6");
                salaryRange.NumberFormat = "¥#,##0.00";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ImportFromArray_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功从数组导入数据: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 从数组导入数据时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 导出数据到数组示例
        /// 演示如何从Excel工作表导出数据到内存数组
        /// </summary>
        static void ExportToArrayExample()
        {
            Console.WriteLine("=== 导出数据到数组示例 ===");

            try
            {
                // 首先创建一个包含数据的工作簿
                CreateSampleDataFile();

                // 打开包含数据的工作簿
                using var excelApp = ExcelFactory.Open("SampleData.xlsx");

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;

                // 导出数据到数组
                var dataRange = worksheet.Range("A1:F6");
                object[,] exportedData = (object[,])dataRange.Value;

                // 显示导出的数据
                Console.WriteLine("从Excel导出的数据:");
                for (int row = 0; row < exportedData.GetLength(0); row++)
                {
                    for (int col = 0; col < exportedData.GetLength(1); col++)
                    {
                        Console.Write($"{exportedData[row, col]}\t");
                    }
                    Console.WriteLine();
                }

                // 保存工作簿副本
                string fileName = $"ExportToArray_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功导出数据到数组: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 导出数据到数组时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 从CSV文件导入数据示例
        /// 演示如何从CSV文件导入数据到Excel
        /// </summary>
        static void ImportFromCsvExample()
        {
            Console.WriteLine("=== 从CSV文件导入数据示例 ===");

            try
            {
                // 首先创建一个CSV文件
                CreateSampleCsvFile();

                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "CSV导入数据";

                // 打开CSV文件
                using var csvApp = ExcelFactory.Open("SampleData.csv");
                var csvWorkbook = csvApp.ActiveWorkbook;
                var csvWorksheet = csvWorkbook.ActiveSheetWrap;

                // 复制CSV数据到当前工作表
                var csvRange = csvWorksheet.UsedRange;
                var targetRange = worksheet.Range("A1");
                csvRange.Copy(targetRange);

                // 设置标题格式
                var headerRange = worksheet.Range("A1:F1");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGreen;
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ImportFromCsv_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功从CSV文件导入数据: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 从CSV文件导入数据时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 导出数据到CSV文件示例
        /// 演示如何将Excel数据导出到CSV文件
        /// </summary>
        static void ExportToCsvExample()
        {
            Console.WriteLine("=== 导出数据到CSV文件示例 ===");

            try
            {
                // 首先创建一个包含数据的工作簿
                CreateSampleDataFile();

                // 打开包含数据的工作簿
                using var excelApp = ExcelFactory.Open("SampleData.xlsx");

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;

                // 保存为CSV格式
                string csvFileName = $"ExportToCsv_{DateTime.Now:yyyyMMddHHmmss}.csv";
                workbook.SaveAs(csvFileName, XlFileFormat.xlCSV);

                Console.WriteLine($"✓ 成功导出数据到CSV文件: {csvFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 导出数据到CSV文件时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据格式转换示例
        /// 演示如何在Excel中进行数据格式转换
        /// </summary>
        static void DataFormatConversionExample()
        {
            Console.WriteLine("=== 数据格式转换示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数据格式转换";

                // 创建示例数据
                worksheet.Range("A1").Value = "原始数据";
                worksheet.Range("A1").Font.Bold = true;

                worksheet.Range("A2").Value = "数字文本";
                worksheet.Range("B2").Value = "12345";

                worksheet.Range("A3").Value = "日期文本";
                worksheet.Range("B3").Value = "2023-01-01";

                worksheet.Range("A4").Value = "货币文本";
                worksheet.Range("B4").Value = "1234.56";

                worksheet.Range("A5").Value = "百分比文本";
                worksheet.Range("B5").Value = "0.1234";

                // 转换数字文本为数字
                worksheet.Range("C2").Formula = "=VALUE(B2)";
                worksheet.Range("D2").Value = "转换为数字";

                // 转换文本为日期
                worksheet.Range("C3").Formula = "=DATEVALUE(B3)";
                worksheet.Range("C3").NumberFormat = "yyyy-mm-dd";
                worksheet.Range("D3").Value = "转换为日期";

                // 转换为货币格式
                worksheet.Range("C4").Formula = "=VALUE(B4)";
                worksheet.Range("C4").NumberFormat = "¥#,##0.00";
                worksheet.Range("D4").Value = "转换为货币";

                // 转换为百分比格式
                worksheet.Range("C5").Formula = "=VALUE(B5)";
                worksheet.Range("C5").NumberFormat = "0.00%";
                worksheet.Range("D5").Value = "转换为百分比";

                // 设置格式
                var headerRange = worksheet.Range("A1:D1");
                headerRange.Interior.Color = Color.LightBlue;

                var labelRange = worksheet.Range("A2:A5");
                labelRange.Font.Bold = true;

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"DataFormatConversion_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数据格式转换: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数据格式转换时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 创建示例数据文件
        /// 用于导出示例
        /// </summary>
        static void CreateSampleDataFile()
        {
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "示例数据";

                // 创建示例数据
                object[,] employeeData = {
                    {"员工ID", "姓名", "部门", "职位", "入职日期", "薪资"},
                    {1001, "张三", "技术部", "软件工程师", new DateTime(2020, 1, 15), 8000},
                    {1002, "李四", "销售部", "销售经理", new DateTime(2019, 3, 22), 12000},
                    {1003, "王五", "市场部", "市场专员", new DateTime(2021, 7, 10), 7000},
                    {1004, "赵六", "人事部", "人事专员", new DateTime(2020, 11, 5), 6500},
                    {1005, "钱七", "财务部", "会计师", new DateTime(2018, 9, 18), 9000}
                };

                // 将数据导入到工作表
                var dataRange = worksheet.Range("A1:F6");
                dataRange.Value = employeeData;

                // 保存工作簿
                workbook.SaveAs("SampleData.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建示例数据文件时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建示例CSV文件
        /// 用于CSV导入示例
        /// </summary>
        static void CreateSampleCsvFile()
        {
            try
            {
                string csvContent = "员工ID,姓名,部门,职位,入职日期,薪资\n" +
                                   "1001,张三,技术部,软件工程师,2020-01-15,8000\n" +
                                   "1002,李四,销售部,销售经理,2019-03-22,12000\n" +
                                   "1003,王五,市场部,市场专员,2021-07-10,7000\n" +
                                   "1004,赵六,人事部,人事专员,2020-11-05,6500\n" +
                                   "1005,钱七,财务部,会计师,2018-09-18,9000";

                File.WriteAllText("SampleData.csv", csvContent);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建示例CSV文件时出错: {ex.Message}");
            }
        }
    }
}