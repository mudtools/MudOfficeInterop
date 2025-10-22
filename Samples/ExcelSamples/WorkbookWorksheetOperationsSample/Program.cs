//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;

namespace WorkbookWorksheetOperationsSample
{
    /// <summary>
    /// 工作簿与工作表操作示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行工作簿和工作表的基本操作
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("工作簿与工作表操作示例");
            Console.WriteLine("======================");
            Console.WriteLine();

            // 演示工作簿基本操作
            WorkbookBasicOperationsExample();

            // 演示工作表基本操作
            WorksheetBasicOperationsExample();

            // 演示工作簿保护功能
            WorkbookProtectionExample();

            // 演示工作表保护功能
            WorksheetProtectionExample();

            // 演示工作簿和工作表的高级操作
            AdvancedOperationsExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 工作簿基本操作示例
        /// 演示如何创建、检查和操作工作簿
        /// </summary>
        static void WorkbookBasicOperationsExample()
        {
            Console.WriteLine("=== 工作簿基本操作示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取工作簿集合
                var workbooks = excelApp.Workbooks;

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 显示工作簿基本信息
                Console.WriteLine($"工作簿名称: {workbook.Name}");
                Console.WriteLine($"工作簿完整路径: {workbook.FullName}");
                Console.WriteLine($"是否只读: {workbook.ReadOnly}");
                Console.WriteLine($"是否受密码保护: {workbook.HasPassword}");

                // 设置工作簿属性
                workbook.Keywords = "Excel自动化,工作簿操作,示例";
                workbook.Subject = "工作簿操作示例";
                workbook.Author = "MudTools.OfficeInterop.Excel示例";

                // 添加一些数据
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "基本信息";
                worksheet.Range("A1").Value = "工作簿操作示例";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Font.Size = 16;

                worksheet.Range("A3").Value = "这是工作簿基本操作的示例数据";
                worksheet.Range("A4").Value = $"创建时间: {DateTime.Now}";

                // 保存工作簿
                string fileName = $"WorkbookBasicOperations_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建并保存工作簿: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 工作簿基本操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 工作表基本操作示例
        /// 演示如何创建、重命名、删除和管理工作表
        /// </summary>
        static void WorksheetBasicOperationsExample()
        {
            Console.WriteLine("=== 工作表基本操作示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 获取工作表集合
                var worksheets = workbook.Worksheets;

                // 显示初始工作表数量
                Console.WriteLine($"初始工作表数量: {worksheets.Count}");

                // 重命名默认工作表
                var firstWorksheet = workbook.ActiveSheetWrap;
                firstWorksheet.Name = "首页";

                // 添加新工作表
                var newWorksheet = worksheets.Add() as IExcelWorksheet;
                newWorksheet.Name = "数据表";

                // 再添加一个工作表
                var anotherWorksheet = worksheets.Add(after: firstWorksheet);
                anotherWorksheet.Name = "图表";

                // 显示当前工作表数量
                Console.WriteLine($"添加后工作表数量: {worksheets.Count}");

                // 在工作表中添加数据
                firstWorksheet.Range("A1").Value = "首页";
                firstWorksheet.Range("A1").Font.Bold = true;

                newWorksheet.Range("A1").Value = "员工姓名";
                newWorksheet.Range("B1").Value = "部门";
                newWorksheet.Range("C1").Value = "薪资";

                newWorksheet.Range("A2").Value = "张三";
                newWorksheet.Range("B2").Value = "技术部";
                newWorksheet.Range("C2").Value = 8000;

                newWorksheet.Range("A3").Value = "李四";
                newWorksheet.Range("B3").Value = "销售部";
                newWorksheet.Range("C3").Value = 7500;

                // 保存工作簿
                string fileName = $"WorksheetBasicOperations_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建并保存包含多个工作表的工作簿: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 工作表基本操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 工作簿保护示例
        /// 演示如何保护和取消保护工作簿
        /// </summary>
        static void WorkbookProtectionExample()
        {
            Console.WriteLine("=== 工作簿保护示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 添加一些工作表
                var worksheets = workbook.Worksheets;
                worksheets.Add().Name = "数据表1";
                worksheets.Add().Name = "数据表2";

                // 保护工作簿结构
                string password = "MyPassword123";
                workbook.Protect(password, true, true);

                // 检查保护状态
                Console.WriteLine($"结构是否受保护: {workbook.ProtectStructure}");
                Console.WriteLine($"窗口是否受保护: {workbook.ProtectWindows}");

                // 保存工作簿
                string fileName = $"ProtectedWorkbook_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                // 取消保护
                workbook.Unprotect(password);
                Console.WriteLine("工作簿保护已取消");

                // 重新保存
                string unprotectedFileName = $"UnprotectedWorkbook_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(unprotectedFileName);

                Console.WriteLine($"✓ 成功演示工作簿保护功能");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 工作簿保护操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 工作表保护示例
        /// 演示如何保护和取消保护工作表
        /// </summary>
        static void WorksheetProtectionExample()
        {
            Console.WriteLine("=== 工作表保护示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "受保护的工作表";

                // 添加一些数据
                worksheet.Range("A1").Value = "受保护的工作表示例";
                worksheet.Range("A1").Font.Bold = true;

                worksheet.Range("A3").Value = "姓名";
                worksheet.Range("B3").Value = "部门";
                worksheet.Range("C3").Value = "薪资";

                worksheet.Range("A4").Value = "张三";
                worksheet.Range("B4").Value = "技术部";
                worksheet.Range("C4").Value = 8000;

                worksheet.Range("A5").Value = "李四";
                worksheet.Range("B5").Value = "销售部";
                worksheet.Range("C5").Value = 7500;

                // 保护工作表
                string password = "SheetPassword123";
                worksheet.Protect(password);

                // 检查保护状态
                Console.WriteLine($"工作表是否受保护: {worksheet.ProtectContents}");

                // 保存工作簿
                string fileName = $"ProtectedWorksheet_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                // 取消保护
                worksheet.Unprotect(password);
                Console.WriteLine("工作表保护已取消");

                // 重新保存
                string unprotectedFileName = $"UnprotectedWorksheet_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(unprotectedFileName);

                Console.WriteLine($"✓ 成功演示工作表保护功能");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 工作表保护操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 工作簿和工作表高级操作示例
        /// 演示移动、复制工作表等高级操作
        /// </summary>
        static void AdvancedOperationsExample()
        {
            Console.WriteLine("=== 工作簿和工作表高级操作示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 创建多个工作表
                var worksheets = workbook.Worksheets;
                var sheet1 = workbook.ActiveSheetWrap;
                sheet1.Name = "源数据";

                var sheet2 = worksheets.Add();
                sheet2.Name = "处理结果";

                var sheet3 = worksheets.Add();
                sheet3.Name = "图表展示";

                // 在源数据工作表中添加数据
                sheet1.Range("A1").Value = "月份";
                sheet1.Range("B1").Value = "销售额";
                sheet1.Range("C1").Value = "利润";

                string[] months = { "1月", "2月", "3月", "4月", "5月", "6月" };
                Random random = new Random();

                for (int i = 0; i < months.Length; i++)
                {
                    sheet1.Range($"A{i + 2}").Value = months[i];
                    sheet1.Range($"B{i + 2}").Value = random.Next(50000, 100000);
                    sheet1.Range($"C{i + 2}").Value = random.Next(5000, 20000);
                }

                // 复制工作表
                var copiedSheet = sheet1.Copy(after: sheet3);
                copiedSheet.Name = "源数据副本";

                // 移动工作表
                sheet2.Move(before: sheet1);

                // 显示工作表顺序
                Console.WriteLine("工作表顺序:");
                for (int i = 1; i <= worksheets.Count; i++)
                {
                    Console.WriteLine($"  {i}. {worksheets.Item(i).Name}");
                }

                // 保存工作簿
                string fileName = $"AdvancedOperations_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示工作簿和工作表高级操作");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 工作簿和工作表高级操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}