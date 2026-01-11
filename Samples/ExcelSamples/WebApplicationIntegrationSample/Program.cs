//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace WebApplicationIntegrationSample
{
    /// <summary>
    /// Web应用集成示例程序
    /// 演示如何将MudTools.OfficeInterop.Excel集成到Web应用中
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static async Task Main(string[] args)
        {
            Console.WriteLine("Web应用集成示例");
            Console.WriteLine("==============");
            Console.WriteLine();

            // 演示服务器端Excel处理
            await ServerSideExcelProcessingExample();

            // 演示文件上传处理
            await FileUploadProcessingExample();

            // 演示在线预览功能
            OnlinePreviewExample();

            // 演示权限控制
            PermissionControlExample();

            // 演示报表生成服务
            ReportGenerationServiceExample();

            // 演示数据导出API
            DataExportApiExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 服务器端Excel处理示例
        /// 演示如何在服务器端处理Excel文件
        /// </summary>
        static async Task ServerSideExcelProcessingExample()
        {
            Console.WriteLine("=== 服务器端Excel处理示例 ===");

            try
            {
                // 模拟创建一个待处理的Excel文件
                string inputFileName = $"InputFile_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                CreateSampleExcelFile(inputFileName);

                // 创建Excel应用程序实例（模拟服务器端处理）
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 配置服务器端优化设置
                excelApp.DisplayAlerts = false;
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;
                excelApp.Calculation = XlCalculation.xlCalculationManual;

                Console.WriteLine("  正在处理Excel文件...");

                // 打开Excel文件
                var workbook = excelApp.OpenWorkbook(inputFileName);
                var worksheet = workbook.ActiveSheetWrap;

                // 执行处理逻辑
                worksheet.Name = "处理后数据";

                // 添加处理时间戳
                worksheet.Range("A1").Value = "数据处理时间";
                worksheet.Range("B1").Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // 模拟数据处理
                worksheet.Range("A2").Value = "处理状态";
                worksheet.Range("B2").Value = "已完成";

                worksheet.Range("A3").Value = "处理记录数";
                worksheet.Range("B3").Value = 1000;

                // 设置格式
                worksheet.Range("A1:B1").Font.Bold = true;
                worksheet.Range("A1:B3").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();

                // 保存处理结果
                string outputFileName = $"Processed_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(outputFileName);

                Console.WriteLine($"  ✓ 成功处理Excel文件: {inputFileName} -> {outputFileName}");

                // 清理临时文件
                if (File.Exists(inputFileName))
                {
                    File.Delete(inputFileName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ✗ 处理Excel文件时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 文件上传处理示例
        /// 演示如何处理上传的Excel文件
        /// </summary>
        static async Task FileUploadProcessingExample()
        {
            Console.WriteLine("=== 文件上传处理示例 ===");

            try
            {
                // 模拟文件上传过程
                Console.WriteLine("  模拟文件上传...");

                // 创建模拟上传的Excel文件
                string uploadFileName = $"UploadedFile_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                CreateSampleExcelFile(uploadFileName);

                Console.WriteLine($"  文件已上传: {uploadFileName}");

                // 处理上传的文件
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 配置处理环境
                excelApp.DisplayAlerts = false;
                excelApp.ScreenUpdating = false;

                // 打开上传的文件
                var workbook = excelApp.OpenWorkbook(uploadFileName);
                var worksheet = workbook.ActiveSheetWrap;

                // 验证文件格式
                if (worksheet.Range("A1").Value?.ToString() != "产品名称")
                {
                    Console.WriteLine("  警告: 文件格式可能不正确");
                }

                // 数据验证和清洗
                worksheet.Range("E1").Value = "上传时间";
                worksheet.Range("E1").Font.Bold = true;
                worksheet.Range("E2").Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                worksheet.Range("E3").Value = "上传状态";
                worksheet.Range("E3").Font.Bold = true;
                worksheet.Range("F3").Value = "已验证";

                // 添加处理日志
                var logSheet = workbook.Worksheets.Add() as IExcelWorksheet;
                logSheet.Name = "处理日志";
                logSheet.Range("A1").Value = "处理日志";
                logSheet.Range("A1").Font.Bold = true;
                logSheet.Range("A1").Interior.Color = Color.LightBlue;

                logSheet.Range("A2").Value = "上传时间";
                logSheet.Range("B2").Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                logSheet.Range("A3").Value = "文件名";
                logSheet.Range("B3").Value = uploadFileName;

                logSheet.Range("A4").Value = "记录数";
                logSheet.Range("B4").Value = 10; // 模拟数据行数

                logSheet.Columns.AutoFit();

                // 保存处理后的文件
                string processedFileName = $"ProcessedUpload_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(processedFileName);

                Console.WriteLine($"  ✓ 成功处理上传文件: {processedFileName}");

                // 清理临时文件
                if (File.Exists(uploadFileName))
                {
                    File.Delete(uploadFileName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ✗ 处理上传文件时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 在线预览功能示例
        /// 演示如何实现Excel文件的在线预览
        /// </summary>
        static void OnlinePreviewExample()
        {
            Console.WriteLine("=== 在线预览功能示例 ===");

            try
            {
                // 创建用于预览的Excel文件
                using var excelApp = ExcelFactory.BlankWorkbook();
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "预览数据";

                // 创建预览数据
                worksheet.Range("A1").Value = "在线预览示例";
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.Navy;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1:E1").Merge(null);

                worksheet.Range("A2").Value = "功能";
                worksheet.Range("B2").Value = "说明";
                worksheet.Range("C2").Value = "状态";
                worksheet.Range("D2").Value = "示例";

                // 设置表头格式
                var headerRange = worksheet.Range("A2:D2");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 预览功能列表
                object[,] previewFeatures = {
                    {"文件打开", "支持打开各种Excel格式", "已实现", "xlsx, xls"},
                    {"数据展示", "完整显示单元格数据", "已实现", "文本、数字、日期"},
                    {"格式保留", "保持原有格式样式", "已实现", "字体、颜色、边框"},
                    {"图表预览", "支持图表显示", "部分实现", "基础图表类型"},
                    {"公式计算", "支持公式自动计算", "已实现", "常用函数"},
                    {"打印预览", "提供打印预览功能", "计划中", "页面设置"}
                };

                var dataRange = worksheet.Range("A3:D8");
                dataRange.Value = previewFeatures;

                // 设置数据格式
                worksheet.Range("A3:D8").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range("C3:C8").Interior.Color = Color.LightGreen;

                // 添加说明
                worksheet.Range("A10").Value = "预览说明:";
                worksheet.Range("A10").Font.Bold = true;

                worksheet.Range("A11").Value = "1. 在线预览功能允许用户在不下载文件的情况下查看Excel内容";
                worksheet.Range("A12").Value = "2. 预览模式下文件为只读，确保数据安全";
                worksheet.Range("A13").Value = "3. 支持大部分Excel格式和功能";
                worksheet.Range("A14").Value = "4. 预览性能优化，支持大文件快速加载";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存预览文件
                string previewFileName = $"OnlinePreview_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(previewFileName);

                Console.WriteLine($"  ✓ 成功创建预览文件: {previewFileName}");
                Console.WriteLine("  在Web应用中可以通过HTML或PDF方式展示Excel内容");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ✗ 创建预览文件时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 权限控制示例
        /// 演示如何实现Excel文件的权限控制
        /// </summary>
        static void PermissionControlExample()
        {
            Console.WriteLine("=== 权限控制示例 ===");

            try
            {
                // 创建带权限控制的Excel文件
                using var excelApp = ExcelFactory.BlankWorkbook();
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "权限控制";

                // 创建权限控制说明
                worksheet.Range("A1").Value = "Excel权限控制系统";
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.DarkRed;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1:G1").Merge(null);

                worksheet.Range("A2").Value = "权限级别";
                worksheet.Range("B2").Value = "可读";
                worksheet.Range("C2").Value = "可写";
                worksheet.Range("D2").Value = "可编辑结构";
                worksheet.Range("E2").Value = "可打印";
                worksheet.Range("F2").Value = "可另存为";
                worksheet.Range("G2").Value = "可分享";

                // 设置表头格式
                var headerRange = worksheet.Range("A2:G2");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 权限级别定义
                object[,] permissionLevels = {
                    {"访客", "√", "×", "×", "×", "×", "×"},
                    {"普通用户", "√", "√", "×", "√", "×", "×"},
                    {"高级用户", "√", "√", "√", "√", "√", "×"},
                    {"管理员", "√", "√", "√", "√", "√", "√"}
                };

                var dataRange = worksheet.Range("A3:G6");
                dataRange.Value = permissionLevels;

                // 设置数据格式
                worksheet.Range("A3:G6").Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range("B3:G6").HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 添加条件格式
                var yesRange = worksheet.Range("B3:G6");
                var yesCondition = yesRange.FormatConditions.Add(
                    XlFormatConditionType.xlCellValue,
                    XlFormatConditionOperator.xlEqual,
                    "√");
                yesCondition.Interior.Color = Color.LightGreen;

                var noCondition = yesRange.FormatConditions.Add(
                    XlFormatConditionType.xlCellValue,
                    XlFormatConditionOperator.xlEqual,
                    "×");
                noCondition.Interior.Color = Color.LightCoral;

                // 权限控制策略
                worksheet.Range("A8").Value = "权限控制策略:";
                worksheet.Range("A8").Font.Bold = true;

                worksheet.Range("A9").Value = "1. 基于角色的访问控制(RBAC)";
                worksheet.Range("A10").Value = "2. 文件级权限设置";
                worksheet.Range("A11").Value = "3. 工作表级权限设置";
                worksheet.Range("A12").Value = "4. 单元格级权限设置";
                worksheet.Range("A13").Value = "5. 时间限制访问控制";
                worksheet.Range("A14").Value = "6. IP地址访问控制";

                // 实现方式
                worksheet.Range("A16").Value = "实现方式:";
                worksheet.Range("A16").Font.Bold = true;

                worksheet.Range("A17").Value = "1. 通过Excel工作簿保护功能";
                worksheet.Range("A18").Value = "2. 通过Web应用层权限验证";
                worksheet.Range("A19").Value = "3. 通过文件系统权限控制";
                worksheet.Range("A20").Value = "4. 通过数据库权限管理";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存权限控制文件
                string permissionFileName = $"PermissionControl_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(permissionFileName);

                Console.WriteLine($"  ✓ 成功创建权限控制文件: {permissionFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ✗ 创建权限控制文件时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 报表生成服务示例
        /// 演示如何实现报表生成服务
        /// </summary>
        static void ReportGenerationServiceExample()
        {
            Console.WriteLine("=== 报表生成服务示例 ===");

            try
            {
                // 模拟报表生成服务
                Console.WriteLine("  启动报表生成服务...");

                // 创建报表模板
                using var excelApp = ExcelFactory.BlankWorkbook();
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "报表模板";

                // 创建报表模板结构
                worksheet.Range("A1").Value = "动态报表生成服务";
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.DarkGreen;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1:F1").Merge(null);

                // 报表参数区域
                worksheet.Range("A2").Value = "参数名称";
                worksheet.Range("B2").Value = "参数值";
                worksheet.Range("C2").Value = "说明";

                var paramHeaderRange = worksheet.Range("A2:C2");
                paramHeaderRange.Font.Bold = true;
                paramHeaderRange.Interior.Color = Color.LightGray;

                object[,] reportParams = {
                    {"报表类型", "{ReportType}", "报表类型标识"},
                    {"开始日期", "{StartDate}", "报表开始日期"},
                    {"结束日期", "{EndDate}", "报表结束日期"},
                    {"部门", "{Department}", "所属部门"},
                    {"生成时间", "{GenerateTime}", "报表生成时间"}
                };

                var paramRange = worksheet.Range("A3:C7");
                paramRange.Value = reportParams;
                paramRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 数据区域占位符
                worksheet.Range("A9").Value = "数据区域";
                worksheet.Range("A9").Font.Bold = true;
                worksheet.Range("A9").Interior.Color = Color.LightBlue;
                worksheet.Range("A9:F20").Borders.LineStyle = XlLineStyle.xlContinuous;

                worksheet.Range("A10").Value = "{DataArea}";

                // 图表区域
                worksheet.Range("H1").Value = "图表区域";
                worksheet.Range("H1").Font.Bold = true;
                worksheet.Range("H1").Interior.Color = Color.LightYellow;
                worksheet.Range("H1:M15").Borders.LineStyle = XlLineStyle.xlContinuous;

                worksheet.Range("H2").Value = "{ChartArea}";

                // 页脚区域
                worksheet.Range("A22").Value = "制表人: {Creator}";
                worksheet.Range("D22").Value = "审核人: {Reviewer}";
                worksheet.Range("A23").Value = "生成时间: {GenerateTime}";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存报表模板
                string templateFileName = $"ReportTemplate_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(templateFileName);

                Console.WriteLine($"  ✓ 成功创建报表模板: {templateFileName}");

                // 模拟基于模板生成实际报表
                using var reportApp = ExcelFactory.CreateFrom(templateFileName);
                var reportWorkbook = reportApp.ActiveWorkbook;
                var reportSheet = reportWorkbook.ActiveSheetWrap;

                // 填充报表数据
                reportSheet.Range("B3").Value = "销售报表";
                reportSheet.Range("B4").Value = "2023-01-01";
                reportSheet.Range("B5").Value = "2023-12-31";
                reportSheet.Range("B6").Value = "所有部门";
                reportSheet.Range("B7").Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // 创建示例数据
                reportSheet.Range("A12").Value = "销售数据示例";
                reportSheet.Range("A12").Font.Bold = true;

                reportSheet.Range("A13").Value = "月份";
                reportSheet.Range("B13").Value = "销售额";
                reportSheet.Range("C13").Value = "成本";
                reportSheet.Range("D13").Value = "利润";

                var headerRange = reportSheet.Range("A13:D13");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;

                object[,] salesData = {
                    {"1月", 100000, 70000, 30000},
                    {"2月", 120000, 80000, 40000},
                    {"3月", 140000, 90000, 50000},
                    {"总计", 360000, 240000, 120000}
                };

                var dataRange = reportSheet.Range("A14:D17");
                dataRange.Value = salesData;
                dataRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 设置数字格式
                reportSheet.Range("B14:D17").NumberFormat = "¥#,##0";

                // 保存生成的报表
                string reportFileName = $"GeneratedReport_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                reportWorkbook.SaveAs(reportFileName);

                Console.WriteLine($"  ✓ 成功生成报表: {reportFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ✗ 报表生成服务出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数据导出API示例
        /// 演示如何实现数据导出功能
        /// </summary>
        static void DataExportApiExample()
        {
            Console.WriteLine("=== 数据导出API示例 ===");

            try
            {
                // 模拟数据导出API调用
                Console.WriteLine("  调用数据导出API...");

                // 创建导出数据
                using var excelApp = ExcelFactory.BlankWorkbook();
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "导出数据";

                // 创建导出API说明
                worksheet.Range("A1").Value = "数据导出API";
                worksheet.Range("A1").Font.Size = 16;
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Interior.Color = Color.Purple;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1:G1").Merge(null);

                worksheet.Range("A2").Value = "API端点";
                worksheet.Range("B2").Value = "/api/export/excel";

                worksheet.Range("A3").Value = "请求方法";
                worksheet.Range("B3").Value = "POST";

                worksheet.Range("A4").Value = "请求参数";
                worksheet.Range("B4").Value = "query, format, filters";

                worksheet.Range("A5").Value = "支持格式";
                worksheet.Range("B5").Value = "xlsx, xls, csv, pdf";

                worksheet.Range("A6").Value = "认证方式";
                worksheet.Range("B6").Value = "Bearer Token";

                // 设置格式
                var apiInfoRange = worksheet.Range("A2:B6");
                apiInfoRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                var apiHeaderRange = worksheet.Range("A2:A6");
                apiHeaderRange.Font.Bold = true;
                apiHeaderRange.Interior.Color = Color.LightGray;

                // 导出数据示例
                worksheet.Range("A8").Value = "导出数据示例";
                worksheet.Range("A8").Font.Bold = true;

                // 创建示例数据
                worksheet.Range("A9").Value = "员工ID";
                worksheet.Range("B9").Value = "姓名";
                worksheet.Range("C9").Value = "部门";
                worksheet.Range("D9").Value = "职位";
                worksheet.Range("E9").Value = "入职日期";
                worksheet.Range("F9").Value = "薪资";

                var dataHeaderRange = worksheet.Range("A9:F9");
                dataHeaderRange.Font.Bold = true;
                dataHeaderRange.Interior.Color = Color.LightBlue;
                dataHeaderRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                object[,] employeeData = {
                    {"E001", "张三", "技术部", "软件工程师", "2020-01-15", 12000},
                    {"E002", "李四", "销售部", "销售经理", "2019-03-20", 15000},
                    {"E003", "王五", "人事部", "人事专员", "2021-07-10", 8000},
                    {"E004", "赵六", "财务部", "会计师", "2018-11-05", 10000},
                    {"E005", "钱七", "市场部", "市场专员", "2022-02-28", 9000}
                };

                var dataRange = worksheet.Range("A10:F14");
                dataRange.Value = employeeData;
                dataRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 设置数字格式
                worksheet.Range("F10:F14").NumberFormat = "¥#,##0";
                worksheet.Range("E10:E14").NumberFormat = "yyyy-mm-dd";

                // 导出选项
                worksheet.Range("A16").Value = "导出选项:";
                worksheet.Range("A16").Font.Bold = true;

                worksheet.Range("A17").Value = "1. 全量导出 - 导出所有数据";
                worksheet.Range("A18").Value = "2. 条件导出 - 根据筛选条件导出";
                worksheet.Range("A19").Value = "3. 分页导出 - 分批次导出大量数据";
                worksheet.Range("A20").Value = "4. 模板导出 - 按指定模板导出";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存导出文件
                string exportFileName = $"DataExport_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(exportFileName);

                Console.WriteLine($"  ✓ 成功创建数据导出示例: {exportFileName}");
                Console.WriteLine("  在Web应用中可通过API接口动态生成并下载Excel文件");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ✗ 数据导出API示例出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 创建示例Excel文件
        /// </summary>
        /// <param name="fileName">文件名</param>
        static void CreateSampleExcelFile(string fileName)
        {
            using var excelApp = ExcelFactory.BlankWorkbook();
            var workbook = excelApp.ActiveWorkbook;
            var worksheet = workbook.ActiveSheetWrap;

            // 创建示例数据
            worksheet.Range("A1").Value = "产品名称";
            worksheet.Range("B1").Value = "单价";
            worksheet.Range("C1").Value = "库存";

            object[,] sampleData = {
                {"笔记本电脑", 5000, 50},
                {"台式电脑", 4000, 30},
                {"平板电脑", 3000, 80},
                {"手机", 2000, 200},
                {"耳机", 500, 500},
                {"键盘", 200, 300},
                {"鼠标", 100, 600},
                {"显示器", 1500, 100},
                {"打印机", 1200, 50},
                {"路由器", 300, 150}
            };

            var dataRange = worksheet.Range("A2:C11");
            dataRange.Value = sampleData;

            // 设置格式
            worksheet.Range("A1:C1").Font.Bold = true;
            worksheet.Range("A1:C1").Interior.Color = Color.LightGray;
            worksheet.Range("A1:C11").Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Columns.AutoFit();

            // 保存文件
            workbook.SaveAs(fileName);
        }
    }
}