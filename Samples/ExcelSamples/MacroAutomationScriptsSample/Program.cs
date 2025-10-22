//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using System.Drawing;

namespace MacroAutomationScriptsSample
{
    /// <summary>
    /// 宏与自动化脚本示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel执行宏和自动化脚本
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("宏与自动化脚本示例");
            Console.WriteLine("===============");
            Console.WriteLine();

            // 演示Excel 4.0宏执行
            Excel4MacroExample();

            // 演示VBA宏执行
            VbaMacroExample();

            // 演示宏安全管理
            MacroSecurityExample();

            // 演示自动化脚本执行
            AutomationScriptExample();

            // 演示宏模块管理
            MacroModuleManagementExample();

            // 演示宏错误处理
            MacroErrorHandlingExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// Excel 4.0宏执行示例
        /// 演示如何执行Excel 4.0宏函数
        /// </summary>
        static void Excel4MacroExample()
        {
            Console.WriteLine("=== Excel 4.0宏执行示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "Excel4宏";

                // 创建示例数据
                worksheet.Range["A1"].Value = "数值1";
                worksheet.Range["B1"].Value = "数值2";
                worksheet.Range["C1"].Value = "求和结果";

                worksheet.Range["A2"].Value = 10;
                worksheet.Range["B2"].Value = 20;

                // 执行Excel 4.0宏函数计算求和
                // 使用SUM函数计算A2和B2的和
                string macroCode = "SUM(A2:B2)";
                object result = excelApp.ExecuteExcel4Macro(macroCode);

                worksheet.Range["C2"].Value = result;

                worksheet.Range["A4"].Value = "Excel 4.0宏执行结果";
                worksheet.Range["A4"].Font.Bold = true;
                worksheet.Range["A4"].Interior.Color = Color.LightBlue;

                // 设置数字格式
                worksheet.Range["A2:C2"].NumberFormat = "0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"Excel4Macro_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示Excel 4.0宏执行: {fileName}");
                Console.WriteLine($"  计算结果: {result}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Excel 4.0宏执行时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// VBA宏执行示例
        /// 演示如何执行VBA宏
        /// </summary>
        static void VbaMacroExample()
        {
            Console.WriteLine("=== VBA宏执行示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 添加VBA模块
                var vbProject = workbook.VBProject;
                var vbComponent = vbProject.VBComponents.Add(MsVBIDE.vbext_ComponentType.vbext_ct_StdModule);
                vbComponent.Name = "SampleModule";

                // 添加VBA代码
                string vbaCode = @"
Sub HelloWorld()
    MsgBox ""Hello, World from VBA Macro!""
End Sub

Function CalculateArea(length As Double, width As Double) As Double
    CalculateArea = length * width
End Function

Sub FormatCurrentSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ws.Range(""A1:C10"").Font.Bold = True
    ws.Range(""A1:C10"").Interior.Color = RGB(200, 200, 200)
    ws.Columns.AutoFit
End Sub
";

                vbComponent.CodeModule.AddFromString(vbaCode);

                // 执行VBA宏
                try
                {
                    // 执行HelloWorld宏（会显示消息框）
                    // 注意：在实际应用中，这会弹出一个消息框
                    // excelApp.Run("HelloWorld");

                    // 执行CalculateArea函数
                    object areaResult = excelApp.Run("CalculateArea", 10.5, 20.3);

                    // 获取活动工作表
                    var worksheet = workbook.ActiveSheetWrap;
                    worksheet.Name = "VBA宏";

                    // 在工作表中显示结果
                    worksheet.Range["A1"].Value = "长度";
                    worksheet.Range["B1"].Value = "宽度";
                    worksheet.Range["C1"].Value = "面积";

                    worksheet.Range["A2"].Value = 10.5;
                    worksheet.Range["B2"].Value = 20.3;
                    worksheet.Range["C2"].Value = areaResult;

                    worksheet.Range["A4"].Value = "VBA宏执行结果";
                    worksheet.Range["A4"].Font.Bold = true;
                    worksheet.Range["A4"].Interior.Color = Color.LightGreen;

                    // 设置数字格式
                    worksheet.Range["A2:C2"].NumberFormat = "0.00";

                    // 自动调整列宽
                    worksheet.Columns.AutoFit();

                    // 保存工作簿
                    string fileName = $"VbaMacro_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    workbook.SaveAs(fileName);

                    Console.WriteLine($"✓ 成功演示VBA宏执行: {fileName}");
                    Console.WriteLine($"  面积计算结果: {areaResult}");
                    Console.WriteLine("  注意：HelloWorld宏会显示消息框，在实际执行时会弹出对话框");
                }
                catch (Exception runEx)
                {
                    Console.WriteLine($"  VBA宏执行警告: {runEx.Message}");
                    Console.WriteLine("  注意：某些宏可能需要启用宏功能才能正常执行");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ VBA宏执行时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 宏安全管理示例
        /// 演示如何管理宏安全设置
        /// </summary>
        static void MacroSecurityExample()
        {
            Console.WriteLine("=== 宏安全管理示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "宏安全";

                // 显示当前宏安全设置信息
                worksheet.Range["A1"].Value = "宏安全设置信息";
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = Color.LightBlue;

                worksheet.Range["A3"].Value = "注意：";
                worksheet.Range["A4"].Value = "1. 宏安全设置通常在Excel选项中配置";
                worksheet.Range["A5"].Value = "2. 不同的安全级别影响宏的执行";
                worksheet.Range["A6"].Value = "3. 企业环境中通常禁用所有宏";
                worksheet.Range["A7"].Value = "4. 可以通过数字签名提高宏可信度";

                // 模拟宏安全级别检查
                worksheet.Range["C3"].Value = "宏安全级别";
                worksheet.Range["C4"].Value = "中等"; // 模拟值

                worksheet.Range["C6"].Value = "建议操作";
                worksheet.Range["C7"].Value = "启用宏前请确认来源可信";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"MacroSecurity_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示宏安全管理: {fileName}");
                Console.WriteLine("  宏安全设置信息已记录到工作表中");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 宏安全管理时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 自动化脚本执行示例
        /// 演示如何执行复杂的自动化脚本
        /// </summary>
        static void AutomationScriptExample()
        {
            Console.WriteLine("=== 自动化脚本执行示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;

                // 添加VBA模块用于自动化脚本
                var vbProject = workbook.VBProject;
                var vbComponent = vbProject.VBComponents.Add(MsVBIDE.vbext_ComponentType.vbext_ct_StdModule);
                vbComponent.Name = "AutomationModule";

                // 添加复杂自动化脚本
                string automationScript = @"
Sub ProcessSalesData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalSales As Double
    Dim avgSales As Double
    
    ' 创建工作表
    Set ws = ActiveSheet
    ws.Name = ""销售数据""
    
    ' 创建表头
    ws.Range(""A1"").Value = ""日期""
    ws.Range(""B1"").Value = ""产品""
    ws.Range(""C1"").Value = ""数量""
    ws.Range(""D1"").Value = ""单价""
    ws.Range(""E1"").Value = ""金额""
    
    ' 添加示例数据
    For i = 1 To 100
        ws.Cells(i + 1, 1).Value = DateAdd(""d"", i, Date)
        ws.Cells(i + 1, 2).Value = ""产品"" & ((i Mod 5) + 1)
        ws.Cells(i + 1, 3).Value = Int(Rnd() * 100) + 1
        ws.Cells(i + 1, 4).Value = Int(Rnd() * 1000) + 100
        ws.Cells(i + 1, 5).Formula = ""=C"" & (i + 1) & ""*D"" & (i + 1)
    Next i
    
    ' 计算统计信息
    lastRow = ws.Cells(ws.Rows.Count, ""A"").End(xlUp).Row
    totalSales = Application.WorksheetFunction.Sum(ws.Range(""E2:E"" & lastRow))
    avgSales = Application.WorksheetFunction.Average(ws.Range(""E2:E"" & lastRow))
    
    ' 输出统计结果
    ws.Cells(lastRow + 2, 1).Value = ""总计""
    ws.Cells(lastRow + 2, 5).Value = totalSales
    ws.Cells(lastRow + 3, 1).Value = ""平均""
    ws.Cells(lastRow + 3, 5).Value = avgSales
    
    ' 格式化数据
    ws.Range(""A1:E1"").Font.Bold = True
    ws.Range(""A1:E1"").Interior.Color = RGB(200, 200, 220)
    ws.Range(""A"" & (lastRow + 2) & "":E"" & (lastRow + 3)).Font.Bold = True
    ws.Columns.AutoFit
End Sub

Sub GenerateReport()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ws.Name = ""自动化报告""
    
    ws.Range(""A1"").Value = ""自动化报告""
    ws.Range(""A1"").Font.Size = 16
    ws.Range(""A1"").Font.Bold = True
    
    ws.Range(""A3"").Value = ""报告生成时间:""
    ws.Range(""B3"").Value = Now()
    
    ws.Range(""A5"").Value = ""状态:""
    ws.Range(""B5"").Value = ""完成""
    
    ws.Columns.AutoFit
End Sub
";

                vbComponent.CodeModule.AddFromString(automationScript);

                // 执行自动化脚本
                try
                {
                    // 执行销售数据处理脚本
                    excelApp.Run("ProcessSalesData");

                    // 添加新工作表并执行报告生成脚本
                    var reportSheet = workbook.Worksheets.Add();
                    excelApp.Run("GenerateReport");

                    // 自动调整列宽
                    workbook.ActiveSheetWrap.Columns.AutoFit();

                    // 保存工作簿
                    string fileName = $"AutomationScript_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    workbook.SaveAs(fileName);

                    Console.WriteLine($"✓ 成功演示自动化脚本执行: {fileName}");
                    Console.WriteLine("  已执行ProcessSalesData和GenerateReport两个自动化脚本");
                }
                catch (Exception runEx)
                {
                    Console.WriteLine($"  自动化脚本执行警告: {runEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 自动化脚本执行时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 宏模块管理示例
        /// 演示如何管理VBA宏模块
        /// </summary>
        static void MacroModuleManagementExample()
        {
            Console.WriteLine("=== 宏模块管理示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "宏模块";

                // 显示宏模块管理信息
                worksheet.Range["A1"].Value = "宏模块管理";
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = Color.LightYellow;

                worksheet.Range["A3"].Value = "VBA项目名称:";
                worksheet.Range["B3"].Value = workbook.VBProject.Name;

                worksheet.Range["A4"].Value = "模块数量:";
                worksheet.Range["B4"].Value = workbook.VBProject.VBComponents.Count;

                // 列出模块信息
                worksheet.Range["A6"].Value = "模块列表:";
                worksheet.Range["A6"].Font.Bold = true;

                for (int i = 1; i <= workbook.VBProject.VBComponents.Count; i++)
                {
                    var component = workbook.VBProject.VBComponents.Item(i);
                    worksheet.Range[$"A{6 + i}"].Value = $"模块 {i}:";
                    worksheet.Range[$"B{6 + i}"].Value = component.Name;
                    worksheet.Range[$"C{6 + i}"].Value = component.Type.ToString();
                }

                worksheet.Range["A10"].Value = "操作说明:";
                worksheet.Range["A11"].Value = "1. 可以添加、删除和修改VBA模块";
                worksheet.Range["A12"].Value = "2. 每个模块可以包含多个过程和函数";
                worksheet.Range["A13"].Value = "3. 模块名称应具有描述性";
                worksheet.Range["A14"].Value = "4. 建议按功能对模块进行分组";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"MacroModuleManagement_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示宏模块管理: {fileName}");
                Console.WriteLine($"  VBA项目包含 {workbook.VBProject.VBComponents.Count} 个模块");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 宏模块管理时出错: {ex.Message}");
                Console.WriteLine("  注意：可能需要在Excel中启用访问VBA项目对象模型");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 宏错误处理示例
        /// 演示如何处理宏执行中的错误
        /// </summary>
        static void MacroErrorHandlingExample()
        {
            Console.WriteLine("=== 宏错误处理示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "错误处理";

                // 创建错误处理示例数据
                worksheet.Range["A1"].Value = "宏错误处理示例";
                worksheet.Range["A1"].Font.Bold = true;
                worksheet.Range["A1"].Interior.Color = Color.LightCoral;

                worksheet.Range["A3"].Value = "常见错误类型:";
                worksheet.Range["A4"].Value = "1. 宏名称不存在";
                worksheet.Range["A5"].Value = "2. 参数类型不匹配";
                worksheet.Range["A6"].Value = "3. 宏执行过程中出错";
                worksheet.Range["A7"].Value = "4. 安全设置阻止宏执行";

                worksheet.Range["A9"].Value = "错误处理策略:";
                worksheet.Range["A10"].Value = "1. 使用Try-Catch捕获异常";
                worksheet.Range["A11"].Value = "2. 验证宏名称和参数";
                worksheet.Range["A12"].Value = "3. 检查宏安全设置";
                worksheet.Range["A13"].Value = "4. 提供友好的错误信息";
                worksheet.Range["A14"].Value = "5. 记录错误日志便于调试";

                worksheet.Range["A16"].Value = "最佳实践:";
                worksheet.Range["A17"].Value = "1. 在执行前验证宏存在";
                worksheet.Range["A18"].Value = "2. 使用有意义的宏名称";
                worksheet.Range["A19"].Value = "3. 提供详细的错误描述";
                worksheet.Range["A20"].Value = "4. 实现重试机制";
                worksheet.Range["A21"].Value = "5. 优雅地处理异常情况";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"MacroErrorHandling_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示宏错误处理: {fileName}");

                // 演示错误处理代码
                Console.WriteLine("  错误处理代码示例:");
                Console.WriteLine("  try {");
                Console.WriteLine("      excelApp.Run(\"NonExistentMacro\");");
                Console.WriteLine("  } catch (Exception ex) {");
                Console.WriteLine("      Console.WriteLine($\"宏执行失败: {ex.Message}\");");
                Console.WriteLine("  }");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 宏错误处理时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}