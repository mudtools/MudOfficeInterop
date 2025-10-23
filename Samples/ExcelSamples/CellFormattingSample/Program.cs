//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace CellFormattingSample
{
    /// <summary>
    /// 单元格格式设置示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行单元格格式设置
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("单元格格式设置示例");
            Console.WriteLine("==================");
            Console.WriteLine();

            // 演示基础字体格式设置
            BasicFontFormattingExample();

            // 演示高级字体格式设置
            AdvancedFontFormattingExample();

            // 演示背景和填充格式设置
            BackgroundAndFillFormattingExample();

            // 演示边框格式设置
            BorderFormattingExample();

            // 演示数字格式设置
            NumberFormattingExample();

            // 演示对齐格式设置
            AlignmentFormattingExample();

            // 演示综合格式设置示例
            ComprehensiveFormattingExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 基础字体格式设置示例
        /// 演示如何设置字体名称、大小、粗体、斜体等基本属性
        /// </summary>
        static void BasicFontFormattingExample()
        {
            Console.WriteLine("=== 基础字体格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "基础字体格式";

                // 设置标题字体格式
                var titleRange = worksheet.Range("A1");
                titleRange.Value = "基础字体格式示例";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 16;
                titleRange.Font.Bold = true;
                titleRange.Font.Color = Color.DarkBlue;
                titleRange.Interior.Color = Color.LightGray;

                // 设置普通文本字体格式
                var normalTextRange = worksheet.Range("A3");
                normalTextRange.Value = "普通文本";
                normalTextRange.Font.Name = "宋体";
                normalTextRange.Font.Size = 10;

                // 设置粗体文本
                var boldTextRange = worksheet.Range("A4");
                boldTextRange.Value = "粗体文本";
                boldTextRange.Font.Bold = true;

                // 设置斜体文本
                var italicTextRange = worksheet.Range("A5");
                italicTextRange.Value = "斜体文本";
                italicTextRange.Font.Italic = true;

                // 设置下划线文本
                var underlineTextRange = worksheet.Range("A6");
                underlineTextRange.Value = "下划线文本";
                underlineTextRange.Font.Underline = XlUnderlineStyle.xlUnderlineStyleSingle;

                // 设置字体颜色
                var coloredTextRange = worksheet.Range("A7");
                coloredTextRange.Value = "彩色文本";
                coloredTextRange.Font.Color = Color.Red;

                // 保存工作簿
                string fileName = $"BasicFontFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示基础字体格式设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 基础字体格式设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 高级字体格式设置示例
        /// 演示如何设置字体的高级属性，如删除线、上标、下标等
        /// </summary>
        static void AdvancedFontFormattingExample()
        {
            Console.WriteLine("=== 高级字体格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "高级字体格式";

                // 设置带删除线的文本
                var strikethroughRange = worksheet.Range("A1");
                strikethroughRange.Value = "带删除线的文本";
                strikethroughRange.Font.Strikethrough = true;

                // 设置上标文本
                var superscriptRange = worksheet.Range("A2");
                superscriptRange.Value = "上标文本";
                superscriptRange.Font.Superscript = true;

                // 设置下标文本
                var subscriptRange = worksheet.Range("A3");
                subscriptRange.Value = "下标文本";
                subscriptRange.Font.Subscript = true;

                // 设置字体颜色索引
                var colorIndexRange = worksheet.Range("A4");
                colorIndexRange.Value = "颜色索引文本";
                colorIndexRange.Font.ColorIndex = 3; // 红色

                // 组合字体样式
                var combinedStyleRange = worksheet.Range("A5");
                combinedStyleRange.Value = "组合样式文本";
                combinedStyleRange.Font.Name = "楷体";
                combinedStyleRange.Font.Size = 14;
                combinedStyleRange.Font.Bold = true;
                combinedStyleRange.Font.Italic = true;
                combinedStyleRange.Font.Underline = XlUnderlineStyle.xlUnderlineStyleDouble;
                combinedStyleRange.Font.Color = Color.Purple;

                // 保存工作簿
                string fileName = $"AdvancedFontFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示高级字体格式设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 高级字体格式设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 背景和填充格式设置示例
        /// 演示如何设置单元格的背景色和填充效果
        /// </summary>
        static void BackgroundAndFillFormattingExample()
        {
            Console.WriteLine("=== 背景和填充格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "背景和填充格式";

                // 设置纯色背景
                var solidColorRange = worksheet.Range("A1");
                solidColorRange.Value = "纯色背景";
                solidColorRange.Interior.Color = Color.LightBlue;

                // 设置颜色索引背景
                var colorIndexRange = worksheet.Range("A2");
                colorIndexRange.Value = "颜色索引背景";
                colorIndexRange.Interior.ColorIndex = XlColorIndex.xlColorIndexAutomatic; // 黄色

                // 设置图案填充
                var patternRange = worksheet.Range("A3");
                patternRange.Value = "图案填充";
                patternRange.Interior.Pattern = XlPattern.xlPatternGray75;
                patternRange.Interior.PatternColor = Color.Gray;

                // 设置渐变填充（如果支持）
                var gradientRange = worksheet.Range("A4");
                gradientRange.Value = "渐变填充";
                gradientRange.Interior.Color = Color.LightGreen;

                // 保存工作簿
                string fileName = $"BackgroundAndFillFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示背景和填充格式设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 背景和填充格式设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 边框格式设置示例
        /// 演示如何设置单元格的边框样式
        /// </summary>
        static void BorderFormattingExample()
        {
            Console.WriteLine("=== 边框格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "边框格式";

                // 设置细线边框
                var thinBorderRange = worksheet.Range("A1");
                thinBorderRange.Value = "细线边框";
                thinBorderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                thinBorderRange.Borders.Weight = XlBorderWeight.xlThin;

                // 设置粗线边框
                var thickBorderRange = worksheet.Range("A2");
                thickBorderRange.Value = "粗线边框";
                thickBorderRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                thickBorderRange.Borders.Weight = XlBorderWeight.xlThick;

                // 设置虚线边框
                var dashBorderRange = worksheet.Range("A3");
                dashBorderRange.Value = "虚线边框";
                dashBorderRange.Borders.LineStyle = XlLineStyle.xlDash;

                // 设置点线边框
                var dotBorderRange = worksheet.Range("A4");
                dotBorderRange.Value = "点线边框";
                dotBorderRange.Borders.LineStyle = XlLineStyle.xlDot;

                // 设置特定边框
                var specificBorderRange = worksheet.Range("A5");
                specificBorderRange.Value = "特定边框";
                specificBorderRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                specificBorderRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                specificBorderRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDash;
                specificBorderRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDash;

                // 保存工作簿
                string fileName = $"BorderFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示边框格式设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 边框格式设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数字格式设置示例
        /// 演示如何设置数字、日期、货币等格式
        /// </summary>
        static void NumberFormattingExample()
        {
            Console.WriteLine("=== 数字格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数字格式";

                // 设置整数格式
                var integerRange = worksheet.Range("A1");
                integerRange.Value = 1234;
                integerRange.NumberFormat = "0";

                // 设置小数格式
                var decimalRange = worksheet.Range("A2");
                decimalRange.Value = 1234.567;
                decimalRange.NumberFormat = "0.00";

                // 设置百分比格式
                var percentageRange = worksheet.Range("A3");
                percentageRange.Value = 0.1234;
                percentageRange.NumberFormat = "0.00%";

                // 设置货币格式
                var currencyRange = worksheet.Range("A4");
                currencyRange.Value = 1234.56;
                currencyRange.NumberFormat = "¥#,##0.00";

                // 设置日期格式
                var dateRange = worksheet.Range("A5");
                dateRange.Value = DateTime.Now;
                dateRange.NumberFormat = "yyyy-mm-dd";

                // 设置时间格式
                var timeRange = worksheet.Range("A6");
                timeRange.Value = DateTime.Now;
                timeRange.NumberFormat = "hh:mm:ss";

                // 设置科学计数法格式
                var scientificRange = worksheet.Range("A7");
                scientificRange.Value = 123456789;
                scientificRange.NumberFormat = "0.00E+00";

                // 保存工作簿
                string fileName = $"NumberFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数字格式设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数字格式设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 对齐格式设置示例
        /// 演示如何设置单元格内容的对齐方式
        /// </summary>
        static void AlignmentFormattingExample()
        {
            Console.WriteLine("=== 对齐格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "对齐格式";

                // 设置水平居中对齐
                var centerRange = worksheet.Range("A1");
                centerRange.Value = "水平居中";
                centerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                centerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                centerRange.Interior.Color = Color.LightBlue;

                // 设置左对齐
                var leftRange = worksheet.Range("A2");
                leftRange.Value = "左对齐";
                leftRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                leftRange.Interior.Color = Color.LightGreen;

                // 设置右对齐
                var rightRange = worksheet.Range("A3");
                rightRange.Value = "右对齐";
                rightRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rightRange.Interior.Color = Color.LightYellow;

                // 设置垂直居中对齐
                var verticalCenterRange = worksheet.Range("B1:B3");
                verticalCenterRange.Value = "垂直居中";
                verticalCenterRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

                // 设置文本自动换行
                var wrapTextRange = worksheet.Range("C1");
                wrapTextRange.Value = "这是一个很长的文本，用于演示自动换行功能";
                wrapTextRange.WrapText = true;
                wrapTextRange.ColumnWidth = 15;
                wrapTextRange.Interior.Color = Color.LightPink;

                // 设置文本方向
                var orientationRange = worksheet.Range("C2");
                orientationRange.Value = "文本方向";
                orientationRange.Orientation = XlOrientation.xlHorizontal;
                orientationRange.Interior.Color = Color.LightGray;

                // 保存工作簿
                string fileName = $"AlignmentFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示对齐格式设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 对齐格式设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 综合格式设置示例
        /// 演示如何综合应用各种格式设置创建专业报表
        /// </summary>
        static void ComprehensiveFormattingExample()
        {
            Console.WriteLine("=== 综合格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "综合格式示例";

                // 创建销售报表标题
                var titleRange = worksheet.Range("A1:E1");
                titleRange.Value = "2023年销售业绩报表";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 18;
                titleRange.Font.Bold = true;
                titleRange.Font.Color = Color.White;
                titleRange.Interior.Color = Color.DarkBlue;
                titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                titleRange.Merge(); // 合并单元格

                // 创建表头
                var headerRange = worksheet.Range("A2:E2");
                string[] headers = { "部门", "销售额", "成本", "利润", "利润率" };
                headerRange.Value = headers;
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 创建数据行
                string[,] data = {
                    { "技术部", "500000", "300000", "200000", "0.4" },
                    { "销售部", "800000", "500000", "300000", "0.375" },
                    { "市场部", "300000", "200000", "100000", "0.333" },
                    { "人事部", "200000", "150000", "50000", "0.25" },
                    { "财务部", "150000", "100000", "50000", "0.333" }
                };

                var dataRange = worksheet.Range("A3:E7");
                dataRange.Value = data;
                dataRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                // 设置数字格式
                worksheet.Range("B3:B7").NumberFormat = "¥#,##0"; // 销售额
                worksheet.Range("C3:C7").NumberFormat = "¥#,##0"; // 成本
                worksheet.Range("D3:D7").NumberFormat = "¥#,##0"; // 利润
                worksheet.Range("E3:E7").NumberFormat = "0.00%";  // 利润率

                // 设置数据对齐
                worksheet.Range("A3:A7").HorizontalAlignment = XlHAlign.xlHAlignLeft;    // 部门左对齐
                worksheet.Range("B3:E7").HorizontalAlignment = XlHAlign.xlHAlignRight;   // 数值右对齐

                // 设置总计行
                var totalRow = worksheet.Range("A8:E8");
                totalRow.Value = new object[,] { { "总计", "=SUM(B3:B7)", "=SUM(C3:C7)", "=SUM(D3:D7)", "=D8/B8" } };
                totalRow.Font.Bold = true;
                totalRow.Interior.Color = Color.LightBlue;
                totalRow.Borders.LineStyle = XlLineStyle.xlContinuous;
                totalRow.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThick;

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ComprehensiveFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示综合格式设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 综合格式设置时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}