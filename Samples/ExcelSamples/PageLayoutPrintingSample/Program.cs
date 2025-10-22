//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace PageLayoutPrintingSample
{
    /// <summary>
    /// 页面布局与打印设置示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行页面布局和打印设置操作
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("页面布局与打印设置示例");
            Console.WriteLine("====================");
            Console.WriteLine();

            // 演示页面方向与纸张大小设置
            PageOrientationAndPaperSizeExample();

            // 演示页边距设置
            PageMarginsExample();

            // 演示页眉页脚设置
            HeaderFooterExample();

            // 演示打印区域设置
            PrintAreaExample();

            // 演示分页符设置
            PageBreakExample();

            // 演示打印预览功能
            PrintPreviewExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 页面方向与纸张大小设置示例
        /// 演示如何设置页面方向和纸张大小
        /// </summary>
        static void PageOrientationAndPaperSizeExample()
        {
            Console.WriteLine("=== 页面方向与纸张大小设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "页面设置";

                // 创建示例数据
                worksheet.Range["A1"].Value = "产品名称";
                worksheet.Range["B1"].Value = "销售数量";
                worksheet.Range["C1"].Value = "单价";
                worksheet.Range["D1"].Value = "总金额";

                object[,] salesData = {
                    {"笔记本电脑", 10, 5000, 50000},
                    {"台式电脑", 5, 4000, 20000},
                    {"平板电脑", 8, 3000, 24000},
                    {"手机", 20, 2000, 40000},
                    {"耳机", 30, 500, 15000},
                    {"键盘", 25, 200, 5000},
                    {"鼠标", 40, 100, 4000},
                    {"显示器", 6, 1500, 9000},
                    {"打印机", 3, 1200, 3600},
                    {"路由器", 15, 300, 4500}
                };

                var dataRange = worksheet.Range["A2:D11"];
                dataRange.Value = salesData;

                // 获取页面设置对象
                var pageSetup = worksheet.PageSetup;

                // 设置A4纵向布局
                pageSetup.Orientation = XlPageOrientation.xlPortrait;
                pageSetup.PaperSize = XlPaperSize.xlPaperA4;
                pageSetup.Zoom = 100;

                worksheet.Range["F1"].Value = "A4纵向布局";
                worksheet.Range["F1"].Font.Bold = true;
                worksheet.Range["F1"].Interior.Color = Color.LightBlue;

                // 复制数据到F列以展示不同布局
                worksheet.Range["A1:D11"].Copy(worksheet.Range["F2"]);

                // 设置A4横向布局
                pageSetup.Orientation = XlPageOrientation.xlLandscape;
                pageSetup.PaperSize = XlPaperSize.xlPaperA4;
                pageSetup.Zoom = 100;

                worksheet.Range["K1"].Value = "A4横向布局";
                worksheet.Range["K1"].Font.Bold = true;
                worksheet.Range["K1"].Interior.Color = Color.LightGreen;

                // 复制数据到K列以展示不同布局
                worksheet.Range["A1:D11"].Copy(worksheet.Range["K2"]);

                // 设置自定义纸张大小
                pageSetup.PaperSize = XlPaperSize.xlPaperLetter;
                pageSetup.Orientation = XlPageOrientation.xlPortrait;

                worksheet.Range["P1"].Value = "Letter纸张";
                worksheet.Range["P1"].Font.Bold = true;
                worksheet.Range["P1"].Interior.Color = Color.LightYellow;

                // 复制数据到P列以展示不同布局
                worksheet.Range["A1:D11"].Copy(worksheet.Range["P2"]);

                // 设置数字格式
                worksheet.Range["B2:D11"].NumberFormat = "0";
                worksheet.Range["G2:I11"].NumberFormat = "0";
                worksheet.Range["L2:N11"].NumberFormat = "0";
                worksheet.Range["Q2:S11"].NumberFormat = "0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"PageOrientationAndPaperSize_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示页面方向与纸张大小设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 页面方向与纸张大小设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 页边距设置示例
        /// 演示如何设置页面边距
        /// </summary>
        static void PageMarginsExample()
        {
            Console.WriteLine("=== 页边距设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "页边距";

                // 创建示例数据
                worksheet.Range["A1"].Value = "月份";
                worksheet.Range["B1"].Value = "销售额";
                worksheet.Range["C1"].Value = "成本";
                worksheet.Range["D1"].Value = "利润";

                object[,] financialData = {
                    {"1月", 100000, 70000, 30000},
                    {"2月", 120000, 80000, 40000},
                    {"3月", 140000, 90000, 50000},
                    {"4月", 130000, 85000, 45000},
                    {"5月", 150000, 95000, 55000},
                    {"6月", 160000, 100000, 60000}
                };

                var dataRange = worksheet.Range["A2:D7"];
                dataRange.Value = financialData;

                // 获取页面设置对象
                var pageSetup = worksheet.PageSetup;

                // 设置标准边距
                pageSetup.LeftMargin = 0.75;
                pageSetup.RightMargin = 0.75;
                pageSetup.TopMargin = 1.0;
                pageSetup.BottomMargin = 1.0;
                pageSetup.HeaderMargin = 0.5;
                pageSetup.FooterMargin = 0.5;

                worksheet.Range["F1"].Value = "标准边距设置";
                worksheet.Range["F1"].Font.Bold = true;
                worksheet.Range["F1"].Interior.Color = Color.LightBlue;

                // 复制数据到F列以展示不同边距
                worksheet.Range["A1:D7"].Copy(worksheet.Range["F2"]);

                // 设置窄边距
                pageSetup.LeftMargin = 0.25;
                pageSetup.RightMargin = 0.25;
                pageSetup.TopMargin = 0.75;
                pageSetup.BottomMargin = 0.75;

                worksheet.Range["K1"].Value = "窄边距设置";
                worksheet.Range["K1"].Font.Bold = true;
                worksheet.Range["K1"].Interior.Color = Color.LightGreen;

                // 复制数据到K列以展示不同边距
                worksheet.Range["A1:D7"].Copy(worksheet.Range["K2"]);

                // 设置宽边距
                pageSetup.LeftMargin = 1.5;
                pageSetup.RightMargin = 1.5;
                pageSetup.TopMargin = 2.0;
                pageSetup.BottomMargin = 2.0;

                worksheet.Range["P1"].Value = "宽边距设置";
                worksheet.Range["P1"].Font.Bold = true;
                worksheet.Range["P1"].Interior.Color = Color.LightYellow;

                // 复制数据到P列以展示不同边距
                worksheet.Range["A1:D7"].Copy(worksheet.Range["P2"]);

                // 设置数字格式
                worksheet.Range["B2:D7"].NumberFormat = "¥#,##0";
                worksheet.Range["G2:I7"].NumberFormat = "¥#,##0";
                worksheet.Range["L2:N7"].NumberFormat = "¥#,##0";
                worksheet.Range["Q2:S7"].NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"PageMargins_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示页边距设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 页边距设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 页眉页脚设置示例
        /// 演示如何设置页眉和页脚
        /// </summary>
        static void HeaderFooterExample()
        {
            Console.WriteLine("=== 页眉页脚设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "页眉页脚";

                // 创建示例数据
                worksheet.Range["A1"].Value = "产品类别";
                worksheet.Range["B1"].Value = "产品名称";
                worksheet.Range["C1"].Value = "销售数量";
                worksheet.Range["D1"].Value = "销售金额";

                object[,] salesData = {
                    {"电子产品", "笔记本电脑", 10, 50000},
                    {"电子产品", "手机", 20, 60000},
                    {"家居用品", "沙发", 5, 15000},
                    {"家居用品", "床", 3, 9000},
                    {"服装", "T恤", 50, 2500},
                    {"服装", "牛仔裤", 30, 3000}
                };

                var dataRange = worksheet.Range["A2:D7"];
                dataRange.Value = salesData;

                // 获取页面设置对象
                var pageSetup = worksheet.PageSetup;

                // 设置简单页眉页脚
                pageSetup.LeftHeader = "销售报表";
                pageSetup.CenterHeader = "第 &P 页，共 &N 页";
                pageSetup.RightHeader = "打印日期: &D";

                pageSetup.LeftFooter = "公司名称";
                pageSetup.CenterFooter = "机密文件";
                pageSetup.RightFooter = "页码: &P/&N";

                worksheet.Range["F1"].Value = "带页眉页脚的报表";
                worksheet.Range["F1"].Font.Bold = true;
                worksheet.Range["F1"].Interior.Color = Color.LightBlue;

                // 复制数据到F列以展示页眉页脚
                worksheet.Range["A1:D7"].Copy(worksheet.Range["F2"]);

                // 设置复杂页眉页脚
                pageSetup.LeftHeader = "&\"微软雅黑,Bold\"&14销售分析报表";
                pageSetup.CenterHeader = "&A";
                pageSetup.RightHeader = "打印时间: &T";

                pageSetup.LeftFooter = "&Z&F";
                pageSetup.CenterFooter = "&\"宋体\"&10第 &P 页 / 共 &N 页";
                pageSetup.RightFooter = "&D";

                worksheet.Range["K1"].Value = "复杂页眉页脚";
                worksheet.Range["K1"].Font.Bold = true;
                worksheet.Range["K1"].Interior.Color = Color.LightGreen;

                // 复制数据到K列以展示页眉页脚
                worksheet.Range["A1:D7"].Copy(worksheet.Range["K2"]);

                // 设置数字格式
                worksheet.Range["C2:D7"].NumberFormat = "0";
                worksheet.Range["H2:I7"].NumberFormat = "0";
                worksheet.Range["M2:N7"].NumberFormat = "0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"HeaderFooter_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示页眉页脚设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 页眉页脚设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 打印区域设置示例
        /// 演示如何设置打印区域
        /// </summary>
        static void PrintAreaExample()
        {
            Console.WriteLine("=== 打印区域设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "打印区域";

                // 创建示例数据
                worksheet.Range["A1"].Value = "产品名称";
                worksheet.Range["B1"].Value = "1月";
                worksheet.Range["C1"].Value = "2月";
                worksheet.Range["D1"].Value = "3月";
                worksheet.Range["E1"].Value = "4月";
                worksheet.Range["F1"].Value = "5月";
                worksheet.Range["G1"].Value = "6月";

                object[,] salesData = {
                    {"笔记本电脑", 10000, 12000, 14000, 13000, 15000, 16000},
                    {"台式电脑", 8000, 9000, 10000, 9500, 11000, 12000},
                    {"平板电脑", 6000, 7000, 8000, 7500, 9000, 10000},
                    {"手机", 20000, 22000, 24000, 23000, 25000, 26000},
                    {"耳机", 5000, 6000, 7000, 6500, 8000, 9000}
                };

                var dataRange = worksheet.Range["A2:G6"];
                dataRange.Value = salesData;

                // 设置打印区域
                worksheet.PageSetup.PrintArea = "$A$1:$D$6";

                worksheet.Range["I1"].Value = "设置打印区域为A1:D6";
                worksheet.Range["I1"].Font.Bold = true;
                worksheet.Range["I1"].Interior.Color = Color.LightBlue;

                // 复制数据到I列以展示打印区域
                worksheet.Range["A1:G6"].Copy(worksheet.Range["I2"]);

                // 更改打印区域
                worksheet.PageSetup.PrintArea = "$A$1:$G$6";

                worksheet.Range["A8"].Value = "完整数据表格";
                worksheet.Range["A8"].Font.Bold = true;
                worksheet.Range["A8"].Interior.Color = Color.LightGreen;

                // 设置数字格式
                worksheet.Range["B2:G6"].NumberFormat = "0";
                worksheet.Range["J2:O6"].NumberFormat = "0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"PrintArea_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示打印区域设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 打印区域设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 分页符设置示例
        /// 演示如何设置分页符
        /// </summary>
        static void PageBreakExample()
        {
            Console.WriteLine("=== 分页符设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "分页符";

                // 创建大量示例数据
                worksheet.Range["A1"].Value = "序号";
                worksheet.Range["B1"].Value = "产品名称";
                worksheet.Range["C1"].Value = "销售数量";
                worksheet.Range["D1"].Value = "销售金额";

                // 生成50行数据
                for (int i = 1; i <= 50; i++)
                {
                    worksheet.Range[$"A{i + 1}"].Value = i;
                    worksheet.Range[$"B{i + 1}"].Value = $"产品{i}";
                    worksheet.Range[$"C{i + 1}"].Value = i * 10;
                    worksheet.Range[$"D{i + 1}"].Value = i * 1000;
                }

                // 设置水平分页符
                worksheet.HPageBreaks.Add(worksheet.Range["A11"]); // 第10行后分页
                worksheet.HPageBreaks.Add(worksheet.Range["A21"]); // 第20行后分页
                worksheet.HPageBreaks.Add(worksheet.Range["A31"]); // 第30行后分页
                worksheet.HPageBreaks.Add(worksheet.Range["A41"]); // 第40行后分页

                // 设置垂直分页符
                worksheet.VPageBreaks.Add(worksheet.Range["C1"]); // C列前分页

                worksheet.Range["F1"].Value = "分页符设置示例";
                worksheet.Range["F1"].Font.Bold = true;
                worksheet.Range["F1"].Interior.Color = Color.LightBlue;

                // 设置数字格式
                worksheet.Range["A2:A51"].NumberFormat = "0";
                worksheet.Range["C2:C51"].NumberFormat = "0";
                worksheet.Range["D2:D51"].NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"PageBreak_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示分页符设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 分页符设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 打印预览功能示例
        /// 演示如何进行打印预览
        /// </summary>
        static void PrintPreviewExample()
        {
            Console.WriteLine("=== 打印预览功能示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "打印预览";

                // 创建示例数据
                worksheet.Range["A1"].Value = "部门";
                worksheet.Range["B1"].Value = "员工姓名";
                worksheet.Range["C1"].Value = "基本工资";
                worksheet.Range["D1"].Value = "绩效奖金";
                worksheet.Range["E1"].Value = "总工资";

                object[,] employeeData = {
                    {"销售部", "张三", 8000, 2000, 10000},
                    {"销售部", "李四", 8500, 2500, 11000},
                    {"技术部", "王五", 10000, 1500, 11500},
                    {"技术部", "赵六", 11000, 1000, 12000},
                    {"人事部", "钱七", 7000, 1000, 8000},
                    {"财务部", "孙八", 9000, 1200, 10200}
                };

                var dataRange = worksheet.Range["A2:E7"];
                dataRange.Value = employeeData;

                // 设置页面布局
                var pageSetup = worksheet.PageSetup;
                pageSetup.Orientation = XlPageOrientation.xlPortrait;
                pageSetup.PaperSize = XlPaperSize.xlPaperA4;

                // 设置页眉页脚
                pageSetup.CenterHeader = "员工工资表";
                pageSetup.CenterFooter = "第 &P 页";

                // 设置打印标题行
                worksheet.PageSetup.PrintTitleRows = "$1:$1";

                // 显示打印预览（注意：在实际应用中，这会打开Excel的打印预览窗口）
                // worksheet.PrintPreview();

                worksheet.Range["A9"].Value = "打印预览设置已完成";
                worksheet.Range["A9"].Font.Bold = true;
                worksheet.Range["A9"].Interior.Color = Color.LightGreen;

                // 设置数字格式
                worksheet.Range["C2:E7"].NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"PrintPreview_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示打印预览功能: {fileName}");
                Console.WriteLine("注意：在实际应用中，PrintPreview()方法会打开Excel的打印预览窗口");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 打印预览功能时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}