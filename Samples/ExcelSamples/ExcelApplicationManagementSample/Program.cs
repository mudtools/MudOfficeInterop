//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;

namespace ExcelApplicationManagementSample
{
    /// <summary>
    /// Excel应用程序管理示例程序
    /// 演示如何使用ExcelFactory类的各种方法创建和管理Excel应用程序实例
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("Excel应用程序管理示例");
            Console.WriteLine("======================");
            Console.WriteLine();

            // 演示创建空白工作簿
            BlankWorkbookExample();

            // 演示通过ProgID创建特定版本实例
            CreateInstanceExample();

            // 演示基于模板创建工作簿
            CreateFromTemplateExample();

            // 演示打开现有工作簿
            OpenExistingWorkbookExample();

            // 演示连接到现有Excel实例
            ConnectionExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 创建空白工作簿示例
        /// 演示如何使用ExcelFactory.BlankWorkbook()方法创建Excel应用程序实例
        /// </summary>
        static void BlankWorkbookExample()
        {
            Console.WriteLine("=== 创建空白工作簿示例 ===");

            try
            {
                // 使用BlankWorkbook方法创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 获取活动工作表
                var worksheet = workbook.ActiveSheetWrap;

                // 设置工作表名称
                worksheet.Name = "空白工作簿示例";

                // 添加标题
                worksheet.Range("A1").Value = "Excel应用程序管理 - 空白工作簿示例";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Interior.Color = System.Drawing.Color.LightBlue;

                // 添加示例数据
                worksheet.Range("A3").Value = "这是使用ExcelFactory.BlankWorkbook()方法创建的空白工作簿";
                worksheet.Range("A4").Value = "该方法会启动Excel应用程序并创建一个包含一个工作表的空白工作簿";
                worksheet.Range("A5").Value = "创建时间: " + DateTime.Now.ToString();

                // 保存工作簿
                string fileName = $"BlankWorkbookExample_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建空白工作簿: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 创建空白工作簿时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 通过ProgID创建特定版本实例示例
        /// 演示如何使用ExcelFactory.CreateInstance()方法创建特定版本的Excel应用程序实例
        /// </summary>
        static void CreateInstanceExample()
        {
            Console.WriteLine("=== 通过ProgID创建特定版本实例示例 ===");

            try
            {
                // 使用CreateInstance方法创建特定版本的Excel应用程序实例
                // Excel.Application是Excel的ProgID
                using var excelApp = ExcelFactory.CreateInstance("Excel.Application");

                // 设置Excel应用程序可见性
                excelApp.Visible = true;

                // 创建新工作簿
                var workbook = excelApp.BlankWorkbook();

                // 获取活动工作表
                var worksheet = workbook.ActiveSheetWrap;

                // 设置工作表名称
                worksheet.Name = "特定版本示例";

                // 添加标题
                worksheet.Range("A1").Value = "Excel应用程序管理 - 特定版本实例示例";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Interior.Color = System.Drawing.Color.LightGreen;

                // 添加示例数据
                worksheet.Range("A3").Value = "这是使用ExcelFactory.CreateInstance()方法创建的特定版本Excel实例";
                worksheet.Range("A4").Value = "通过指定ProgID可以创建特定版本的Excel应用程序";
                worksheet.Range("A5").Value = "Excel版本: " + excelApp.Version;
                worksheet.Range("A6").Value = "创建时间: " + DateTime.Now.ToString();

                // 保存工作簿
                string fileName = $"CreateInstanceExample_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功创建特定版本实例工作簿: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 通过ProgID创建特定版本实例时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 基于模板创建工作簿示例
        /// 演示如何使用ExcelFactory.CreateFrom()方法基于模板创建工作簿
        /// </summary>
        static void CreateFromTemplateExample()
        {
            Console.WriteLine("=== 基于模板创建工作簿示例 ===");

            try
            {
                // 首先创建一个模板文件
                CreateTemplateFile();

                // 使用CreateFrom方法基于模板创建工作簿
                using var excelApp = ExcelFactory.CreateFrom("Template.xlsx");

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 获取活动工作表
                var worksheet = workbook.ActiveSheetWrap;

                // 填充模板中的数据
                worksheet.Range("B2").Value = "张三";
                worksheet.Range("B3").Value = "Excel自动化工程师";
                worksheet.Range("B4").Value = DateTime.Now.ToString("yyyy-MM-dd");
                worksheet.Range("B5").Value = "北京市朝阳区";

                // 保存工作簿
                string fileName = $"CreateFromTemplateExample_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功基于模板创建工作簿: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 基于模板创建工作簿时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 打开现有工作簿示例
        /// 演示如何使用ExcelFactory.Open()方法打开现有工作簿
        /// </summary>
        static void OpenExistingWorkbookExample()
        {
            Console.WriteLine("=== 打开现有工作簿示例 ===");

            try
            {
                // 首先创建一个要打开的文件
                CreateExistingFile();

                // 使用Open方法打开现有工作簿
                using var excelApp = ExcelFactory.Open("ExistingFile.xlsx");

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 获取活动工作表
                var worksheet = workbook.ActiveSheetWrap;

                // 读取现有数据
                var existingValue = worksheet.Range("A1").Value;
                Console.WriteLine($"原始数据: {existingValue}");

                // 修改数据
                worksheet.Range("A3").Value = "这是通过ExcelFactory.Open()方法打开并修改的数据";
                worksheet.Range("A4").Value = "修改时间: " + DateTime.Now.ToString();

                // 保存工作簿
                string fileName = $"OpenExistingWorkbookExample_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功打开并修改现有工作簿: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 打开现有工作簿时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 连接到现有Excel实例示例
        /// 演示如何使用ExcelFactory.Connection()方法连接到现有Excel实例
        /// </summary>
        static void ConnectionExample()
        {
            Console.WriteLine("=== 连接到现有Excel实例示例 ===");

            try
            {
                // 启动一个Excel实例用于连接测试
                using var testExcelApp = ExcelFactory.BlankWorkbook();
                testExcelApp.Visible = true; // 使Excel可见以便观察

                // 在实际应用中，这里需要获取现有Excel实例的COM对象
                // 由于演示限制，我们直接使用testExcelApp的COM对象
                var comObj = testExcelApp.ComObject;

                // 使用Connection方法连接到现有Excel实例
                var connectedApp = ExcelFactory.Connection(comObj);

                if (connectedApp != null)
                {
                    // 获取活动工作簿
                    var workbook = connectedApp.ActiveWorkbook;

                    // 获取活动工作表
                    var worksheet = workbook.ActiveSheetWrap;

                    // 添加连接信息
                    worksheet.Range("A7").Value = "这是通过ExcelFactory.Connection()方法连接到现有实例后添加的数据";
                    worksheet.Range("A8").Value = "连接时间: " + DateTime.Now.ToString();

                    Console.WriteLine("✓ 成功连接到现有Excel实例");
                }
                else
                {
                    Console.WriteLine("✗ 无法连接到现有Excel实例");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 连接到现有Excel实例时出错: {ex.Message}");
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
                var worksheet = workbook.ActiveSheetWrap;

                // 设置模板内容
                worksheet.Range("A1").Value = "员工信息表";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Interior.Color = System.Drawing.Color.LightYellow;

                worksheet.Range("A2").Value = "姓名:";
                worksheet.Range("A3").Value = "职位:";
                worksheet.Range("A4").Value = "入职日期:";
                worksheet.Range("A5").Value = "地址:";

                // 添加边框
                worksheet.Range("A1:B6").Borders.LineStyle = XlLineStyle.xlContinuous;

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
        static void CreateExistingFile()
        {
            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;

                // 设置文件内容
                worksheet.Range("A1").Value = "这是要被打开的Excel文件";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = System.Drawing.Color.LightCoral;

                worksheet.Range("A2").Value = "该文件将通过ExcelFactory.Open()方法打开";

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