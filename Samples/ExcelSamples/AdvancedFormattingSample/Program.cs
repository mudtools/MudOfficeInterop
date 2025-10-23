//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求.
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Excel;
using System.Drawing;

namespace AdvancedFormattingSample
{
    /// <summary>
    /// 高级格式化示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行高级格式化操作
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("高级格式化示例");
            Console.WriteLine("==============");
            Console.WriteLine();

            // 演示数据验证功能
            DataValidationExample();

            // 演示条件格式设置
            ConditionalFormattingExample();

            // 演示超链接功能
            HyperlinkExample();

            // 演示注释功能
            CommentExample();

            // 演示合并单元格功能
            MergeCellsExample();

            // 演示综合高级格式化示例
            ComprehensiveAdvancedFormattingExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 数据验证示例
        /// 演示如何设置不同类型的数据验证规则
        /// </summary>
        static void DataValidationExample()
        {
            Console.WriteLine("=== 数据验证示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数据验证";

                // 设置整数验证规则
                var integerRange = worksheet.Range("A1");
                integerRange.Value = "整数验证 (1-100)";
                var integerValidationRange = worksheet.Range("A2");
                var integerValidation = integerValidationRange.Validation;
                integerValidation.Add(XlDVType.xlValidateWholeNumber, XlDVAlertStyle.xlValidAlertStop, "1", "100");
                integerValidation.InputTitle = "整数输入";
                integerValidation.InputMessage = "请输入1到100之间的整数";
                integerValidation.ErrorTitle = "输入错误";
                integerValidation.ErrorMessage = "请输入1到100之间的有效整数";
                integerValidation.ShowInput = true;
                integerValidation.ShowError = true;

                // 设置小数验证规则
                var decimalRange = worksheet.Range("B1");
                decimalRange.Value = "小数验证 (0.0-10.0)";
                var decimalValidationRange = worksheet.Range("B2");
                var decimalValidation = decimalValidationRange.Validation;
                decimalValidation.Add(XlDVType.xlValidateDecimal, XlDVAlertStyle.xlValidAlertStop, "0.0", "10.0");
                decimalValidation.InputTitle = "小数输入";
                decimalValidation.InputMessage = "请输入0.0到10.0之间的小数";
                decimalValidation.ErrorTitle = "输入错误";
                decimalValidation.ErrorMessage = "请输入0.0到10.0之间的有效小数";
                decimalValidation.ShowInput = true;
                decimalValidation.ShowError = true;

                // 设置列表验证规则
                var listRange = worksheet.Range("C1");
                listRange.Value = "列表验证";
                var listValidationRange = worksheet.Range("C2");
                var listValidation = listValidationRange.Validation;
                listValidation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, "选项1,选项2,选项3,选项4", "");
                listValidation.InputTitle = "列表选择";
                listValidation.InputMessage = "请选择列表中的一个选项";
                listValidation.ErrorTitle = "输入错误";
                listValidation.ErrorMessage = "请选择列表中的有效选项";
                listValidation.ShowInput = true;
                listValidation.ShowError = true;

                // 设置日期验证规则
                var dateRange = worksheet.Range("D1");
                dateRange.Value = "日期验证";
                var dateValidationRange = worksheet.Range("D2");
                var dateValidation = dateValidationRange.Validation;
                dateValidation.Add(XlDVType.xlValidateDate, XlDVAlertStyle.xlValidAlertStop, "2020/1/1", "2030/12/31");
                dateValidation.InputTitle = "日期输入";
                dateValidation.InputMessage = "请输入2020年到2030年之间的日期";
                dateValidation.ErrorTitle = "输入错误";
                dateValidation.ErrorMessage = "请输入有效的日期（2020年到2030年之间）";
                dateValidation.ShowInput = true;
                dateValidation.ShowError = true;

                // 设置文本长度验证规则
                var textRange = worksheet.Range("E1");
                textRange.Value = "文本长度验证 (最多10字符)";
                var textValidationRange = worksheet.Range("E2");
                var textValidation = textValidationRange.Validation;
                textValidation.Add(XlDVType.xlValidateTextLength, XlDVAlertStyle.xlValidAlertStop, "", "10");
                textValidation.InputTitle = "文本输入";
                textValidation.InputMessage = "请输入最多10个字符的文本";
                textValidation.ErrorTitle = "输入错误";
                textValidation.ErrorMessage = "文本长度不能超过10个字符";
                textValidation.ShowInput = true;
                textValidation.ShowError = true;

                // 保存工作簿
                string fileName = $"DataValidation_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数据验证功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数据验证示例时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 条件格式设置示例
        /// 演示如何设置条件格式来动态改变单元格外观
        /// </summary>
        static void ConditionalFormattingExample()
        {
            Console.WriteLine("=== 条件格式设置示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "条件格式";

                // 创建示例数据
                worksheet.Range("A1").Value = "销售数据";
                worksheet.Range("A1").Font.Bold = true;

                worksheet.Range("A2").Value = "员工";
                worksheet.Range("B2").Value = "销售额";
                worksheet.Range("C2").Value = "目标";
                worksheet.Range("D2").Value = "完成率";

                string[,] employeeData = {
                    {"张三", "80000", "100000", "=B3/C3"},
                    {"李四", "120000", "100000", "=B4/C4"},
                    {"王五", "90000", "100000", "=B5/C5"},
                    {"赵六", "110000", "100000", "=B6/C6"},
                    {"钱七", "95000", "100000", "=B7/C7"}
                };

                var dataRange = worksheet.Range("A3:D7");
                dataRange.Value = employeeData;

                // 设置表头格式
                var headerRange = worksheet.Range("A2:D2");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;

                // 添加条件格式 - 高于目标的销售额显示为绿色
                var aboveTargetRange = worksheet.Range("D3:D7");
                var aboveTargetFormat = aboveTargetRange.FormatConditions.Add(
                    XlFormatConditionType.xlCellValue,
                    XlFormatConditionOperator.xlGreater,
                    "1");
                aboveTargetFormat.Interior.Color = Color.LightGreen;

                // 添加条件格式 - 低于目标的销售额显示为红色
                var belowTargetFormat = aboveTargetRange.FormatConditions.Add(
                    XlFormatConditionType.xlCellValue,
                    XlFormatConditionOperator.xlLess,
                    "1");
                belowTargetFormat.Interior.Color = Color.LightCoral;

                // 设置数字格式
                worksheet.Range("B3:B7").NumberFormat = "¥#,##0";
                worksheet.Range("C3:C7").NumberFormat = "¥#,##0";
                worksheet.Range("D3:D7").NumberFormat = "0.00%";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ConditionalFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示条件格式设置: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 条件格式设置示例时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 超链接示例
        /// 演示如何在单元格中添加超链接
        /// </summary>
        static void HyperlinkExample()
        {
            Console.WriteLine("=== 超链接示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "超链接";

                // 添加标题
                worksheet.Range("A1").Value = "超链接示例";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Font.Size = 16;

                // 添加网页链接
                worksheet.Range("A3").Value = "访问百度";
                worksheet.Hyperlinks.Add(worksheet.Range("A3"), "https://www.baidu.com");

                // 添加邮箱链接
                worksheet.Range("A4").Value = "发送邮件";
                worksheet.Hyperlinks.Add(worksheet.Range("A4"), "mailto:someone@example.com");

                // 添加文件链接
                worksheet.Range("A5").Value = "打开文件";
                worksheet.Hyperlinks.Add(worksheet.Range("A5"), "Sample.xlsx");

                // 添加工作表内链接
                worksheet.Range("A6").Value = "跳转到A10";
                worksheet.Hyperlinks.Add(worksheet.Range("A6"), "", "A10", "", "跳转到单元格A10");
                worksheet.Range("A10").Value = "这里是A10单元格";
                worksheet.Range("A10").Interior.Color = Color.Yellow;

                // 保存工作簿
                string fileName = $"Hyperlink_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示超链接功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 超链接示例时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 注释示例
        /// 演示如何添加和管理单元格注释
        /// </summary>
        static void CommentExample()
        {
            Console.WriteLine("=== 注释示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "注释";

                // 添加标题
                worksheet.Range("A1").Value = "注释示例";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Font.Size = 16;

                // 添加带注释的单元格
                var cellA3 = worksheet.Range("A3");
                cellA3.Value = "带注释的单元格";
                cellA3.AddComment("这是一个单元格注释\n可以包含多行文本\n用于提供额外信息");
                cellA3.Comment.Visible = true;

                // 添加另一个带注释的单元格
                var cellB5 = worksheet.Range("B5");
                cellB5.Value = "另一个注释";
                cellB5.AddComment("这是另一个注释\n可以用于解释数据来源或计算方法");

                // 保存工作簿
                string fileName = $"Comment_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示注释功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 注释示例时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 合并单元格示例
        /// 演示如何合并和取消合并单元格
        /// </summary>
        static void MergeCellsExample()
        {
            Console.WriteLine("=== 合并单元格示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "合并单元格";

                // 添加标题
                worksheet.Range("A1").Value = "合并单元格示例";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Font.Size = 16;

                // 水平合并单元格
                var horizontalMergeRange = worksheet.Range("A3:E3");
                horizontalMergeRange.Value = "水平合并的单元格";
                horizontalMergeRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                horizontalMergeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                horizontalMergeRange.Merge();
                horizontalMergeRange.Interior.Color = Color.LightBlue;

                // 垂直合并单元格
                var verticalMergeRange = worksheet.Range("A5:A10");
                verticalMergeRange.Value = "垂直合并的单元格";
                verticalMergeRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                verticalMergeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                verticalMergeRange.Merge();
                verticalMergeRange.Interior.Color = Color.LightGreen;
                verticalMergeRange.Orientation = XlOrientation.xlVertical; // 垂直文本

                // 合并后居中
                var centerMergeRange = worksheet.Range("C5:E10");
                centerMergeRange.Value = "合并后居中";
                centerMergeRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                centerMergeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                centerMergeRange.Merge();
                centerMergeRange.Interior.Color = Color.LightYellow;

                // 取消合并示例
                var unmergeRange = worksheet.Range("A12:E12");
                unmergeRange.Value = "这是将要被取消合并的区域";
                unmergeRange.Merge();
                // 取消合并将在下面操作中完成

                // 保存工作簿
                string fileName = $"MergeCells_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示合并单元格功能: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 合并单元格示例时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 综合高级格式化示例
        /// 演示多种高级格式化技术的综合应用
        /// </summary>
        static void ComprehensiveAdvancedFormattingExample()
        {
            Console.WriteLine("=== 综合高级格式化示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "综合高级格式化";

                // 创建员工信息表
                worksheet.Range("A1").Value = "员工信息录入表";
                worksheet.Range("A1").Font.Bold = true;
                worksheet.Range("A1").Font.Size = 18;
                worksheet.Range("A1").Font.Color = Color.White;
                worksheet.Range("A1").Interior.Color = Color.DarkBlue;
                var titleRange = worksheet.Range("A1:E1");
                titleRange.Merge();
                titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 创建表头
                worksheet.Range("A2").Value = "姓名";
                worksheet.Range("B2").Value = "部门";
                worksheet.Range("C2").Value = "年龄";
                worksheet.Range("D2").Value = "工资";
                worksheet.Range("E2").Value = "入职日期";

                var headerRange = worksheet.Range("A2:E2");
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGray;
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // 设置数据验证
                // 姓名列 - 文本长度限制
                var nameValidation = worksheet.Range("A3:A20").Validation;
                nameValidation.Add(XlDVType.xlValidateTextLength, XlDVAlertStyle.xlValidAlertStop, "", "20");
                nameValidation.InputTitle = "姓名输入";
                nameValidation.InputMessage = "请输入员工姓名（最多20个字符）";
                nameValidation.ErrorTitle = "输入错误";
                nameValidation.ErrorMessage = "姓名长度不能超过20个字符";
                nameValidation.ShowInput = true;
                nameValidation.ShowError = true;

                // 部门列 - 列表选择
                var departmentValidation = worksheet.Range("B3:B20").Validation;
                departmentValidation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, "技术部,销售部,市场部,人事部,财务部", "");
                departmentValidation.InputTitle = "部门选择";
                departmentValidation.InputMessage = "请选择员工所属部门";
                departmentValidation.ErrorTitle = "输入错误";
                departmentValidation.ErrorMessage = "请选择有效的部门";
                departmentValidation.ShowInput = true;
                departmentValidation.ShowError = true;

                // 年龄列 - 整数范围
                var ageValidation = worksheet.Range("C3:C20").Validation;
                ageValidation.Add(XlDVType.xlValidateWholeNumber, XlDVAlertStyle.xlValidAlertStop, "18", "65");
                ageValidation.InputTitle = "年龄输入";
                ageValidation.InputMessage = "请输入员工年龄（18-65岁）";
                ageValidation.ErrorTitle = "输入错误";
                ageValidation.ErrorMessage = "年龄必须在18-65岁之间";
                ageValidation.ShowInput = true;
                ageValidation.ShowError = true;

                // 工资列 - 小数范围
                var salaryValidation = worksheet.Range("D3:D20").Validation;
                salaryValidation.Add(XlDVType.xlValidateDecimal, XlDVAlertStyle.xlValidAlertStop, "3000", "100000");
                salaryValidation.InputTitle = "工资输入";
                salaryValidation.InputMessage = "请输入员工工资（3000-100000）";
                salaryValidation.ErrorTitle = "输入错误";
                salaryValidation.ErrorMessage = "工资必须在3000-100000之间";
                salaryValidation.ShowInput = true;
                salaryValidation.ShowError = true;

                // 入职日期列 - 日期范围
                var dateValidation = worksheet.Range("E3:E20").Validation;
                dateValidation.Add(XlDVType.xlValidateDate, XlDVAlertStyle.xlValidAlertStop, "2020/1/1", "2030/12/31");
                dateValidation.InputTitle = "日期输入";
                dateValidation.InputMessage = "请输入入职日期（2020-2030年）";
                dateValidation.ErrorTitle = "输入错误";
                dateValidation.ErrorMessage = "请输入有效的入职日期";
                dateValidation.ShowInput = true;
                dateValidation.ShowError = true;

                // 添加示例数据
                worksheet.Range("A3").Value = "张三";
                worksheet.Range("B3").Value = "技术部";
                worksheet.Range("C3").Value = 28;
                worksheet.Range("D3").Value = 12000;
                worksheet.Range("E3").Value = DateTime.Now.AddYears(-2);

                worksheet.Range("A4").Value = "李四";
                worksheet.Range("B4").Value = "销售部";
                worksheet.Range("C4").Value = 32;
                worksheet.Range("D4").Value = 15000;
                worksheet.Range("E4").Value = DateTime.Now.AddYears(-1);

                // 设置数字格式
                worksheet.Range("C3:C20").NumberFormat = "0"; // 年龄
                worksheet.Range("D3:D20").NumberFormat = "¥#,##0.00"; // 工资
                worksheet.Range("E3:E20").NumberFormat = "yyyy-mm-dd"; // 日期

                // 添加超链接到说明文档
                worksheet.Range("A20").Value = "填写说明";
                worksheet.Hyperlinks.Add(worksheet.Range("A20"), "https://www.example.com/guide", "", "", "点击查看填写说明");

                // 添加注释
                worksheet.Range("C2").AddComment("年龄必须在18-65岁之间");
                worksheet.Range("D2").AddComment("工资单位为人民币元");

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ComprehensiveAdvancedFormatting_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示综合高级格式化: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 综合高级格式化示例时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}