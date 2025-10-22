using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System;

namespace ExcelAutomationIntroductionSample
{
    /// <summary>
    /// Excel自动化入门示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel创建和操作Excel文件
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("欢迎使用Excel自动化入门示例！");
            Console.WriteLine("本示例将演示如何使用MudTools.OfficeInterop.Excel创建Excel文件。");
            Console.WriteLine();

            // 演示创建空白工作簿
            CreateBlankWorkbookExample();
            
            // 演示基于模板创建工作簿
            CreateFromTemplateExample();
            
            // 演示打开现有工作簿
            OpenExistingWorkbookExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 创建空白工作簿示例
        /// 对应文档中的第一个Excel自动化应用示例
        /// </summary>
        static void CreateBlankWorkbookExample()
        {
            Console.WriteLine("=== 创建空白工作簿示例 ===");
            
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();
                
                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;
                
                // 获取活动工作表
                var worksheet = workbook.ActiveSheet;
                
                // 在单元格A1中写入数据
                worksheet.Range["A1"].Value = "Hello, Excel Automation!";
                
                // 设置单元格格式
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Font.Size = 14;
                worksheet.Range["A1"].Font.Color = System.Drawing.Color.Blue;
                
                // 在其他单元格中添加更多数据
                worksheet.Range["A3"].Value = "这是一个简单的Excel自动化示例";
                worksheet.Range["A4"].Value = "展示了如何使用MudTools.OfficeInterop.Excel库";
                worksheet.Range["A5"].Value = "日期: " + DateTime.Now.ToString("yyyy-MM-dd");
                
                // 保存工作簿
                string fileName = $"MyFirstExcelApp_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);
                
                Console.WriteLine($"✓ 成功创建Excel文件: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建空白工作簿时出错: {ex.Message}");
            }
            
            Console.WriteLine();
        }

        /// <summary>
        /// 基于模板创建工作簿示例
        /// 演示如何使用模板创建Excel文件
        /// </summary>
        static void CreateFromTemplateExample()
        {
            Console.WriteLine("=== 基于模板创建工作簿示例 ===");
            
            try
            {
                // 首先创建一个模板文件
                CreateTemplateFile();
                
                // 基于模板创建工作簿
                using var excelApp = ExcelFactory.CreateFrom("Template.xlsx");
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                
                // 修改模板中的数据
                worksheet.Range["B2"].Value = "张三";
                worksheet.Range["B3"].Value = DateTime.Now.ToString("yyyy-MM-dd");
                worksheet.Range["B4"].Value = "Excel自动化工程师";
                
                // 保存新文件
                string fileName = $"EmployeeInfo_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);
                
                Console.WriteLine($"✓ 成功基于模板创建Excel文件: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 基于模板创建工作簿时出错: {ex.Message}");
            }
            
            Console.WriteLine();
        }

        /// <summary>
        /// 打开现有工作簿示例
        /// 演示如何打开和修改现有Excel文件
        /// </summary>
        static void OpenExistingWorkbookExample()
        {
            Console.WriteLine("=== 打开现有工作簿示例 ===");
            
            try
            {
                // 首先创建一个要打开的文件
                CreateFileToOpen();
                
                // 打开现有工作簿
                using var excelApp = ExcelFactory.Open("ExistingFile.xlsx");
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                
                // 读取现有数据
                var existingValue = worksheet.Range["A1"].Value;
                Console.WriteLine($"原始数据: {existingValue}");
                
                // 修改数据
                worksheet.Range["A2"].Value = "这是新增的数据";
                worksheet.Range["A3"].Value = $"修改时间: {DateTime.Now}";
                
                // 保存修改后的文件
                string fileName = $"ModifiedFile_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);
                
                Console.WriteLine($"✓ 成功打开并修改Excel文件，保存为: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 打开现有工作簿时出错: {ex.Message}");
            }
            
            Console.WriteLine();
        }

        /// <summary>
        /// 创建模板文件
        /// 用于CreateFromTemplateExample示例
        /// </summary>
        static void CreateTemplateFile()
        {
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                
                // 设置模板内容
                worksheet.Range["A1"].Value = "员工信息表";
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Font.Size = 16;
                
                worksheet.Range["A2"].Value = "姓名:";
                worksheet.Range["A3"].Value = "入职日期:";
                worksheet.Range["A4"].Value = "职位:";
                
                worksheet.Range["B2"].Value = "[姓名]";
                worksheet.Range["B3"].Value = "[入职日期]";
                worksheet.Range["B4"].Value = "[职位]";
                
                // 添加边框
                worksheet.Range["A1:C5"].Borders.LineStyle = XlLineStyle.xlContinuous;
                
                // 保存为模板文件
                workbook.SaveAs("Template.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建模板文件时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建用于打开示例的文件
        /// 用于OpenExistingWorkbookExample示例
        /// </summary>
        static void CreateFileToOpen()
        {
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();
                
                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheet;
                
                // 设置文件内容
                worksheet.Range["A1"].Value = "这是要被打开的Excel文件";
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = System.Drawing.Color.LightBlue;
                
                // 保存文件
                workbook.SaveAs("ExistingFile.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建待打开文件时出错: {ex.Message}");
            }
        }
    }
}