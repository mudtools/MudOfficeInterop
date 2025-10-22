//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Diagnostics;
using System.Drawing;

namespace PerformanceOptimizationSample
{
    /// <summary>
    /// 性能优化示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行性能优化
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("性能优化示例");
            Console.WriteLine("==========");
            Console.WriteLine();

            // 演示内存管理优化
            MemoryManagementOptimizationExample();

            // 演示批处理操作优化
            BatchOperationOptimizationExample();

            // 演示屏幕更新优化
            ScreenUpdateOptimizationExample();

            // 演示计算优化
            CalculationOptimizationExample();

            // 演示资源释放优化
            ResourceReleaseOptimizationExample();

            // 演示大数据处理优化
            LargeDataProcessingOptimizationExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 内存管理优化示例
        /// 演示如何优化内存使用
        /// </summary>
        static void MemoryManagementOptimizationExample()
        {
            Console.WriteLine("=== 内存管理优化示例 ===");

            try
            {
                // 记录初始内存使用情况
                long initialMemory = GC.GetTotalMemory(false) / (1024 * 1024);
                Console.WriteLine($"初始内存使用: {initialMemory} MB");

                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "内存优化";

                // 创建大量数据以观察内存变化
                worksheet.Range("A1").Value = "ID";
                worksheet.Range("B1").Value = "名称";
                worksheet.Range("C1").Value = "数值";

                // 使用数组方式填充大量数据（优化方法）
                object[,] data = new object[10000, 3];
                for (int i = 0; i < 10000; i++)
                {
                    data[i, 0] = i + 1;
                    data[i, 1] = $"项目{i + 1}";
                    data[i, 2] = (i + 1) * 10;
                }

                worksheet.Range("A2:C10001").Value = data;

                // 强制垃圾回收以获取更准确的内存使用情况
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                long afterMemory = GC.GetTotalMemory(false) / (1024 * 1024);
                Console.WriteLine($"填充数据后内存使用: {afterMemory} MB");

                // 显示内存使用信息
                worksheet.Range("E1").Value = "内存优化示例";
                worksheet.Range("E1").Font.Bold = true;
                worksheet.Range("E1").Interior.Color = Color.LightBlue;

                worksheet.Range("E3").Value = "初始内存使用:";
                worksheet.Range("F3").Value = $"{initialMemory} MB";

                worksheet.Range("E4").Value = "数据填充后内存:";
                worksheet.Range("F4").Value = $"{afterMemory} MB";

                worksheet.Range("E5").Value = "内存增长:";
                worksheet.Range("F5").Value = $"{afterMemory - initialMemory} MB";

                // 设置数字格式
                worksheet.Range("A2:A10001").NumberFormat = "0";
                worksheet.Range("C2:C10001").NumberFormat = "0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"MemoryManagementOptimization_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示内存管理优化: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 内存管理优化时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 批处理操作优化示例
        /// 演示如何通过批处理操作提高性能
        /// </summary>
        static void BatchOperationOptimizationExample()
        {
            Console.WriteLine("=== 批处理操作优化示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "批处理优化";

                // 创建示例数据结构
                worksheet.Range("A1").Value = "产品ID";
                worksheet.Range("B1").Value = "产品名称";
                worksheet.Range("C1").Value = "类别";
                worksheet.Range("D1").Value = "价格";
                worksheet.Range("E1").Value = "库存";

                Stopwatch stopwatch = Stopwatch.StartNew();

                // 低效方法：逐单元格设置（不推荐）
                stopwatch.Restart();
                for (int i = 2; i <= 1001; i++)
                {
                    worksheet.Range($"A{i}").Value = i - 1;
                    worksheet.Range($"B{i}").Value = $"产品{i - 1}";
                    worksheet.Range($"C{i}").Value = $"类别{((i - 1) % 5) + 1}";
                    worksheet.Range($"D{i}").Value = (i - 1) * 10.5;
                    worksheet.Range($"E{i}").Value = (i - 1) * 5;
                }
                stopwatch.Stop();
                long inefficientTime = stopwatch.ElapsedMilliseconds;

                // 添加新工作表进行高效方法演示
                var efficientSheet = workbook.Worksheets.Add() as IExcelWorksheet;
                efficientSheet.Name = "高效批处理";

                efficientSheet.Range("A1").Value = "产品ID";
                efficientSheet.Range("B1").Value = "产品名称";
                efficientSheet.Range("C1").Value = "类别";
                efficientSheet.Range("D1").Value = "价格";
                efficientSheet.Range("E1").Value = "库存";

                // 高效方法：使用数组批量设置（推荐）
                stopwatch.Restart();
                object[,] batchData = new object[1000, 5];
                for (int i = 0; i < 1000; i++)
                {
                    batchData[i, 0] = i + 1;
                    batchData[i, 1] = $"产品{i + 1}";
                    batchData[i, 2] = $"类别{((i + 1) % 5) + 1}";
                    batchData[i, 3] = (i + 1) * 10.5;
                    batchData[i, 4] = (i + 1) * 5;
                }

                efficientSheet.Range("A2:E1001").Value = batchData;
                stopwatch.Stop();
                long efficientTime = stopwatch.ElapsedMilliseconds;

                // 显示性能对比结果
                var resultSheet = workbook.Worksheets.Add() as IExcelWorksheet;
                resultSheet.Name = "性能对比";

                resultSheet.Range("A1").Value = "批处理操作优化示例";
                resultSheet.Range("A1").Font.Bold = true;
                resultSheet.Range("A1").Interior.Color = Color.LightGreen;

                resultSheet.Range("A3").Value = "操作方法";
                resultSheet.Range("B3").Value = "耗时(毫秒)";
                resultSheet.Range("C3").Value = "性能说明";

                resultSheet.Range("A4").Value = "逐单元格设置";
                resultSheet.Range("B4").Value = inefficientTime;
                resultSheet.Range("C4").Value = "低效方法，多次COM调用";

                resultSheet.Range("A5").Value = "数组批量设置";
                resultSheet.Range("B5").Value = efficientTime;
                resultSheet.Range("C5").Value = "高效方法，一次COM调用";

                resultSheet.Range("A7").Value = "性能提升:";
                resultSheet.Range("B7").Formula = $"=ROUND((B4-B5)/B4*100, 2) & \"%\"";

                // 设置数字格式
                efficientSheet.Range("A2:A1001").NumberFormat = "0";
                efficientSheet.Range("D2:D1001").NumberFormat = "0.00";
                efficientSheet.Range("E2:E1001").NumberFormat = "0";
                resultSheet.Range("B4:B5").NumberFormat = "0";

                // 自动调整列宽
                efficientSheet.Columns.AutoFit();
                resultSheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"BatchOperationOptimization_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示批处理操作优化: {fileName}");
                Console.WriteLine($"  逐单元格设置耗时: {inefficientTime} 毫秒");
                Console.WriteLine($"  数组批量设置耗时: {efficientTime} 毫秒");
                Console.WriteLine($"  性能提升: {Math.Round((double)(inefficientTime - efficientTime) / inefficientTime * 100, 2)}%");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 批处理操作优化时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 屏幕更新优化示例
        /// 演示如何通过控制屏幕更新提高性能
        /// </summary>
        static void ScreenUpdateOptimizationExample()
        {
            Console.WriteLine("=== 屏幕更新优化示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "屏幕更新优化";

                // 显示屏幕更新优化说明
                worksheet.Range("A1").Value = "屏幕更新优化示例";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.LightYellow;

                worksheet.Range("A3").Value = "优化原理:";
                worksheet.Range("A4").Value = "1. 屏幕更新会消耗大量资源";
                worksheet.Range("A5").Value = "2. 在大量操作时禁用屏幕更新可提高性能";
                worksheet.Range("A6").Value = "3. 操作完成后重新启用屏幕更新";

                worksheet.Range("A8").Value = "优化方法:";
                worksheet.Range("A9").Value = "1. 设置Application.ScreenUpdating = false";
                worksheet.Range("A10").Value = "2. 执行大量操作";
                worksheet.Range("A11").Value = "3. 设置Application.ScreenUpdating = true";

                worksheet.Range("A13").Value = "注意事项:";
                worksheet.Range("A14").Value = "1. 禁用期间用户看不到操作过程";
                worksheet.Range("A15").Value = "2. 必须确保最终重新启用屏幕更新";
                worksheet.Range("A16").Value = "3. 异常处理中也要重新启用屏幕更新";

                // 演示性能测试
                worksheet.Range("C3").Value = "性能测试:";
                worksheet.Range("C4").Value = "启用屏幕更新操作时间:";

                // 模拟启用屏幕更新的操作
                excelApp.ScreenUpdating = true;
                Stopwatch stopwatch = Stopwatch.StartNew();
                for (int i = 1; i <= 100; i++)
                {
                    worksheet.Range($"C{i + 5}").Value = $"测试数据{i}";
                }
                stopwatch.Stop();
                long timeWithScreenUpdate = stopwatch.ElapsedMilliseconds;

                worksheet.Range("D4").Value = $"{timeWithScreenUpdate} 毫秒";

                // 模拟禁用屏幕更新的操作
                worksheet.Range("C107").Value = "禁用屏幕更新操作时间:";
                excelApp.ScreenUpdating = false;
                stopwatch.Restart();
                for (int i = 1; i <= 100; i++)
                {
                    worksheet.Range($"C{i + 107}").Value = $"测试数据{i}";
                }
                stopwatch.Stop();
                long timeWithoutScreenUpdate = stopwatch.ElapsedMilliseconds;

                // 重新启用屏幕更新
                excelApp.ScreenUpdating = true;

                worksheet.Range("D107").Value = $"{timeWithoutScreenUpdate} 毫秒";

                worksheet.Range("C105").Value = "性能提升:";
                worksheet.Range("D105").Formula = $"=ROUND(ABS(D4-D107)/D4*100, 2) & \"%\"";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ScreenUpdateOptimization_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示屏幕更新优化: {fileName}");
                Console.WriteLine($"  启用屏幕更新耗时: {timeWithScreenUpdate} 毫秒");
                Console.WriteLine($"  禁用屏幕更新耗时: {timeWithoutScreenUpdate} 毫秒");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 屏幕更新优化时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 计算优化示例
        /// 演示如何优化Excel计算性能
        /// </summary>
        static void CalculationOptimizationExample()
        {
            Console.WriteLine("=== 计算优化示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "计算优化";

                // 创建包含公式的示例数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "增长率";
                worksheet.Range("D1").Value = "累计销售额";

                // 填充基础数据
                object[,] salesData = new object[12, 4];
                Random random = new Random();

                for (int i = 0; i < 12; i++)
                {
                    salesData[i, 0] = $"第{i + 1}月";
                    salesData[i, 1] = 10000 + random.Next(5000);

                    if (i == 0)
                    {
                        salesData[i, 2] = 0;
                    }
                    else
                    {
                        salesData[i, 2] = $"=ROUND((B{i + 2}-B{i + 1})/B{i + 1}*100, 2)";
                    }

                    if (i == 0)
                    {
                        salesData[i, 3] = $"=B{i + 2}";
                    }
                    else
                    {
                        salesData[i, 3] = $"=D{i + 1}+B{i + 2}";
                    }
                }

                worksheet.Range("A2:D13").Value = salesData;

                // 演示不同计算模式的性能
                worksheet.Range("F1").Value = "计算优化示例";
                worksheet.Range("F1").Font.Bold = true;
                worksheet.Range("F1").Interior.Color = Color.LightBlue;

                worksheet.Range("F3").Value = "计算模式";
                worksheet.Range("G3").Value = "说明";

                worksheet.Range("F4").Value = "自动计算";
                worksheet.Range("G4").Value = "默认模式，输入时自动计算";

                worksheet.Range("F5").Value = "手动计算";
                worksheet.Range("G5").Value = "手动触发计算，提高批量操作性能";

                worksheet.Range("F6").Value = "计算一次";
                worksheet.Range("G6").Value = "仅计算一次，适用于静态数据";

                // 设置计算模式并测量性能
                worksheet.Range("F8").Value = "性能测试:";

                // 自动计算模式测试
                excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
                worksheet.Range("F9").Value = "自动计算模式:";
                Stopwatch stopwatch = Stopwatch.StartNew();
                worksheet.Calculate();
                stopwatch.Stop();
                worksheet.Range("G9").Value = $"{stopwatch.ElapsedMilliseconds} 毫秒";

                // 手动计算模式测试
                excelApp.Calculation = XlCalculation.xlCalculationManual;
                worksheet.Range("F10").Value = "手动计算模式:";
                stopwatch.Restart();
                worksheet.Calculate();
                stopwatch.Stop();
                worksheet.Range("G10").Value = $"{stopwatch.ElapsedMilliseconds} 毫秒";

                // 恢复自动计算模式
                excelApp.Calculation = XlCalculation.xlCalculationAutomatic;

                // 设置数字格式
                worksheet.Range("B2:B13").NumberFormat = "¥#,##0";
                worksheet.Range("C2:C13").NumberFormat = "0.00%";
                worksheet.Range("D2:D13").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"CalculationOptimization_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示计算优化: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 计算优化时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 资源释放优化示例
        /// 演示如何正确释放COM资源
        /// </summary>
        static void ResourceReleaseOptimizationExample()
        {
            Console.WriteLine("=== 资源释放优化示例 ===");

            long initialMemory = GC.GetTotalMemory(false) / (1024 * 1024);

            try
            {
                // 正确的资源释放方式 - 使用using语句
                IExcelApplication excelApp = null;
                IExcelWorkbook workbook = null;
                IExcelWorksheet worksheet = null;

                try
                {
                    excelApp = ExcelFactory.BlankWorkbook();
                    workbook = excelApp.ActiveWorkbook;
                    worksheet = workbook.ActiveSheetWrap;
                    worksheet.Name = "资源释放";

                    // 执行一些操作
                    worksheet.Range("A1").Value = "资源释放优化示例";
                    worksheet.Range("A1").Font.Bold = true;
                    worksheet.Range("A1").Interior.Color = Color.LightCoral;

                    worksheet.Range("A3").Value = "最佳实践:";
                    worksheet.Range("A4").Value = "1. 使用using语句自动释放资源";
                    worksheet.Range("A5").Value = "2. 显式调用Dispose方法";
                    worksheet.Range("A6").Value = "3. 及时关闭工作簿和应用程序";
                    worksheet.Range("A7").Value = "4. 避免循环引用";

                    worksheet.Range("A9").Value = "错误示例:";
                    worksheet.Range("A10").Value = "1. 忘记释放COM对象";
                    worksheet.Range("A11").Value = "2. 不正确的释放顺序";
                    worksheet.Range("A12").Value = "3. 释放后继续使用对象";

                    // 保存工作簿
                    string fileName = $"ResourceReleaseOptimization_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    workbook.SaveAs(fileName);

                    Console.WriteLine($"✓ 成功演示资源释放优化: {fileName}");
                }
                finally
                {
                    // 显式释放资源
                    worksheet?.Dispose();
                    workbook?.Dispose();
                    excelApp?.Dispose();
                }

                // 强制垃圾回收
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                long finalMemory = GC.GetTotalMemory(false) / (1024 * 1024);
                Console.WriteLine($"  内存变化: {initialMemory} MB -> {finalMemory} MB");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 资源释放优化时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 大数据处理优化示例
        /// 演示如何优化大数据处理性能
        /// </summary>
        static void LargeDataProcessingOptimizationExample()
        {
            Console.WriteLine("=== 大数据处理优化示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "大数据优化";

                // 显示大数据处理优化说明
                worksheet.Range("A1").Value = "大数据处理优化示例";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.LightGreen;

                worksheet.Range("A3").Value = "优化策略:";
                worksheet.Range("A4").Value = "1. 分批处理大量数据";
                worksheet.Range("A5").Value = "2. 使用数组进行批量操作";
                worksheet.Range("A6").Value = "3. 适时禁用屏幕更新";
                worksheet.Range("A7").Value = "4. 合理设置计算模式";
                worksheet.Range("A8").Value = "5. 及时释放临时对象";

                worksheet.Range("A10").Value = "性能测试:";

                // 模拟大数据处理 - 分批处理
                int batchSize = 1000;
                int totalRecords = 5000;
                int batchCount = totalRecords / batchSize;

                worksheet.Range("A12").Value = "处理批次:";
                worksheet.Range("B12").Value = batchCount;

                worksheet.Range("A13").Value = "每批记录数:";
                worksheet.Range("B13").Value = batchSize;

                worksheet.Range("A14").Value = "总记录数:";
                worksheet.Range("B14").Value = totalRecords;

                // 创建表头
                worksheet.Range("D1").Value = "ID";
                worksheet.Range("E1").Value = "数据1";
                worksheet.Range("F1").Value = "数据2";
                worksheet.Range("G1").Value = "计算结果";

                // 分批处理数据
                Stopwatch stopwatch = Stopwatch.StartNew();

                // 禁用屏幕更新以提高性能
                excelApp.ScreenUpdating = false;

                for (int batch = 0; batch < batchCount; batch++)
                {
                    int startRow = batch * batchSize + 2;
                    int endRow = (batch + 1) * batchSize + 1;

                    // 使用数组批量填充数据
                    object[,] batchData = new object[batchSize, 4];
                    for (int i = 0; i < batchSize; i++)
                    {
                        int id = batch * batchSize + i + 1;
                        batchData[i, 0] = id;
                        batchData[i, 1] = $"数据1_{id}";
                        batchData[i, 2] = id * 10;
                        batchData[i, 3] = $"=E{startRow + i}*2";
                    }

                    worksheet.Range($"D{startRow}:G{endRow}").Value = batchData;
                }

                // 重新启用屏幕更新
                excelApp.ScreenUpdating = true;

                stopwatch.Stop();

                worksheet.Range("A16").Value = "处理时间:";
                worksheet.Range("B16").Value = $"{stopwatch.ElapsedMilliseconds} 毫秒";

                worksheet.Range("A18").Value = "优化效果:";
                worksheet.Range("A19").Value = "通过分批处理和数组操作，";
                worksheet.Range("A20").Value = "大大提高了大数据处理性能";

                // 设置数字格式
                worksheet.Range("D2:D5001").NumberFormat = "0";
                worksheet.Range("F2:F5001").NumberFormat = "0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"LargeDataProcessingOptimization_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示大数据处理优化: {fileName}");
                Console.WriteLine($"  处理5000条记录耗时: {stopwatch.ElapsedMilliseconds} 毫秒");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 大数据处理优化时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}