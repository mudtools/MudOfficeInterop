//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using System.Drawing;

namespace FormulaFunctionApplicationSample
{
    /// <summary>
    /// 公式与函数应用示例程序
    /// 演示如何使用MudTools.OfficeInterop.Excel进行公式和函数操作
    /// </summary>
    class Program
    {
        /// <summary>
        /// 程序入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            Console.WriteLine("公式与函数应用示例");
            Console.WriteLine("==================");
            Console.WriteLine();

            // 演示基础公式操作
            BasicFormulaOperationsExample();

            // 演示常用函数应用
            CommonFunctionApplicationExample();

            // 演示条件函数应用
            ConditionalFunctionApplicationExample();

            // 演示查找引用函数应用
            LookupReferenceFunctionApplicationExample();

            // 演示统计函数应用
            StatisticalFunctionApplicationExample();

            // 演示数组公式应用
            ArrayFormulaApplicationExample();

            // 演示公式错误处理
            FormulaErrorHandlingExample();

            Console.WriteLine("所有示例已完成。按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 基础公式操作示例
        /// 演示如何设置和读取基本公式
        /// </summary>
        static void BasicFormulaOperationsExample()
        {
            Console.WriteLine("=== 基础公式操作示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "基础公式";

                // 设置基本公式
                worksheet.Range("A1").Value = "单价";
                worksheet.Range("B1").Value = "数量";
                worksheet.Range("C1").Value = "总价";

                // 设置数据
                worksheet.Range("A2").Value = 10.5;
                worksheet.Range("B2").Value = 5;

                // 设置公式
                worksheet.Range("C2").Formula = "=A2*B2";

                // 读取公式
                var formula = worksheet.Range("C2").Formula;
                var value = worksheet.Range("C2").Value;

                Console.WriteLine($"公式: {formula}");
                Console.WriteLine($"计算结果: {value}");

                // 设置更多数据和公式
                for (int i = 3; i <= 6; i++)
                {
                    worksheet.Range($"A{i}").Value = i * 10;  // 单价
                    worksheet.Range($"B{i}").Value = i;       // 数量
                    worksheet.Range($"C{i}").Formula = $"=A{i}*B{i}";  // 总价
                }

                // 计算总计
                worksheet.Range("A7").Value = "总计";
                worksheet.Range("C7").Formula = "=SUM(C2:C6)";

                // 设置格式
                worksheet.Range("A1:C1").Font.Bold = true;
                worksheet.Range("A1:C1").Interior.Color = Color.LightGray;
                worksheet.Range("A7:C7").Font.Bold = true;
                worksheet.Range("A7:C7").Interior.Color = Color.LightBlue;

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"BasicFormulaOperations_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示基础公式操作: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 基础公式操作时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 常用函数应用示例
        /// 演示SUM、AVERAGE、MAX、MIN等常用函数的使用
        /// </summary>
        static void CommonFunctionApplicationExample()
        {
            Console.WriteLine("=== 常用函数应用示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "常用函数";

                // 创建数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "成本";
                worksheet.Range("D1").Value = "利润";

                string[,] salesData = {
                    {"1月", 50000, 30000},
                    {"2月", 55000, 32000},
                    {"3月", 60000, 35000},
                    {"4月", 58000, 33000},
                    {"5月", 62000, 36000},
                    {"6月", 65000, 38000}
                };

                var dataRange = worksheet.Range("A2:C7");
                dataRange.Value = salesData;

                // 计算利润
                for (int i = 2; i <= 7; i++)
                {
                    worksheet.Range($"D{i}").Formula = $"=B{i}-C{i}";
                }

                // 添加统计行
                int statRow = 9;
                worksheet.Range($"A{statRow}").Value = "统计";
                worksheet.Range($"A{statRow}").Font.Bold = true;

                worksheet.Range($"A{statRow + 1}").Value = "总销售额";
                worksheet.Range($"B{statRow + 1}").Formula = "=SUM(B2:B7)";

                worksheet.Range($"A{statRow + 2}").Value = "平均销售额";
                worksheet.Range($"B{statRow + 2}").Formula = "=AVERAGE(B2:B7)";

                worksheet.Range($"A{statRow + 3}").Value = "最高销售额";
                worksheet.Range($"B{statRow + 3}").Formula = "=MAX(B2:B7)";

                worksheet.Range($"A{statRow + 4}").Value = "最低销售额";
                worksheet.Range($"B{statRow + 4}").Formula = "=MIN(B2:B7)";

                worksheet.Range($"A{statRow + 5}").Value = "总利润";
                worksheet.Range($"B{statRow + 5}").Formula = "=SUM(D2:D7)";

                worksheet.Range($"A{statRow + 6}").Value = "平均利润";
                worksheet.Range($"B{statRow + 6}").Formula = "=AVERAGE(D2:D7)";

                // 设置格式
                worksheet.Range("A1:D1").Font.Bold = true;
                worksheet.Range("A1:D1").Interior.Color = Color.LightGray;
                worksheet.Range($"A{statRow}:B{statRow + 6}").Font.Bold = true;
                worksheet.Range($"A{statRow}:A{statRow + 6}").Interior.Color = Color.LightBlue;

                // 设置数字格式
                worksheet.Range("B2:D7").NumberFormat = "¥#,##0";
                worksheet.Range($"B{statRow + 1}:B{statRow + 6}").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"CommonFunctionApplication_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示常用函数应用: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 常用函数应用时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 条件函数应用示例
        /// 演示IF、AND、OR等条件函数的使用
        /// </summary>
        static void ConditionalFunctionApplicationExample()
        {
            Console.WriteLine("=== 条件函数应用示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "条件函数";

                // 创建员工绩效数据
                worksheet.Range("A1").Value = "员工";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "目标";
                worksheet.Range("D1").Value = "完成率";
                worksheet.Range("E1").Value = "绩效等级";

                object[,] employeeData = {
                    {"张三", 80000, 100000},
                    {"李四", 120000, 100000},
                    {"王五", 90000, 100000},
                    {"赵六", 110000, 100000},
                    {"钱七", 95000, 100000}
                };

                var dataRange = worksheet.Range("A2:C6");
                dataRange.Value = employeeData;

                // 计算完成率
                for (int i = 2; i <= 6; i++)
                {
                    worksheet.Range($"D{i}").Formula = $"=B{i}/C{i}";
                }

                // 根据完成率确定绩效等级
                for (int i = 2; i <= 6; i++)
                {
                    // 使用嵌套IF函数确定绩效等级
                    worksheet.Range($"E{i}").Formula =
                        $"=IF(D{i}>=1.2,\"优秀\",IF(D{i}>=1,\"良好\",IF(D{i}>=0.8,\"合格\",\"待改进\")))";
                }

                // 添加绩效统计
                int statRow = 8;
                worksheet.Range($"A{statRow}").Value = "绩效统计";
                worksheet.Range($"A{statRow}").Font.Bold = true;

                worksheet.Range($"A{statRow + 1}").Value = "优秀人数";
                worksheet.Range($"B{statRow + 1}").Formula = "=COUNTIF(E2:E6,\"优秀\")";

                worksheet.Range($"A{statRow + 2}").Value = "良好人数";
                worksheet.Range($"B{statRow + 2}").Formula = "=COUNTIF(E2:E6,\"良好\")";

                worksheet.Range($"A{statRow + 3}").Value = "合格人数";
                worksheet.Range($"B{statRow + 3}").Formula = "=COUNTIF(E2:E6,\"合格\")";

                worksheet.Range($"A{statRow + 4}").Value = "待改进人数";
                worksheet.Range($"B{statRow + 4}").Formula = "=COUNTIF(E2:E6,\"待改进\")";

                // 设置格式
                worksheet.Range("A1:E1").Font.Bold = true;
                worksheet.Range("A1:E1").Interior.Color = Color.LightGray;

                worksheet.Range($"A{statRow}:B{statRow + 4}").Font.Bold = true;
                worksheet.Range($"A{statRow}:A{statRow + 4}").Interior.Color = Color.LightBlue;

                // 设置数字格式
                worksheet.Range("B2:C6").NumberFormat = "¥#,##0";
                worksheet.Range("D2:D6").NumberFormat = "0.00%";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ConditionalFunctionApplication_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示条件函数应用: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 条件函数应用时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 查找引用函数应用示例
        /// 演示VLOOKUP、HLOOKUP、INDEX、MATCH等查找引用函数的使用
        /// </summary>
        static void LookupReferenceFunctionApplicationExample()
        {
            Console.WriteLine("=== 查找引用函数应用示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "查找引用函数";

                // 创建产品价格表
                worksheet.Range("A1").Value = "产品代码";
                worksheet.Range("B1").Value = "产品名称";
                worksheet.Range("C1").Value = "单价";

                object[,] priceData = {
                    {"P001", "笔记本电脑", 5000},
                    {"P002", "台式电脑", 4000},
                    {"P003", "平板电脑", 2000},
                    {"P004", "手机", 3000},
                    {"P005", "耳机", 200}
                };

                var priceRange = worksheet.Range("A2:C6");
                priceRange.Value = priceData;

                // 创建订单表
                worksheet.Range("E1").Value = "订单号";
                worksheet.Range("F1").Value = "产品代码";
                worksheet.Range("G1").Value = "数量";
                worksheet.Range("H1").Value = "产品名称";
                worksheet.Range("I1").Value = "单价";
                worksheet.Range("J1").Value = "总价";

                object[,] orderData = {
                    {"O001", "P001", 2},
                    {"O002", "P003", 5},
                    {"O003", "P005", 10},
                    {"O004", "P002", 1},
                    {"O005", "P004", 3}
                };

                var orderRange = worksheet.Range("E2:G6");
                orderRange.Value = orderData;

                // 使用VLOOKUP查找产品名称
                for (int i = 2; i <= 6; i++)
                {
                    worksheet.Range($"H{i}").Formula = $"=VLOOKUP(F{i},A:C,2,FALSE)";
                }

                // 使用VLOOKUP查找单价
                for (int i = 2; i <= 6; i++)
                {
                    worksheet.Range($"I{i}").Formula = $"=VLOOKUP(F{i},A:C,3,FALSE)";
                }

                // 计算总价
                for (int i = 2; i <= 6; i++)
                {
                    worksheet.Range($"J{i}").Formula = $"=G{i}*I{i}";
                }

                // 计算订单总金额
                worksheet.Range("E8").Value = "订单总金额";
                worksheet.Range("E8").Font.Bold = true;
                worksheet.Range("J8").Formula = "=SUM(J2:J6)";
                worksheet.Range("J8").Font.Bold = true;

                // 设置格式
                worksheet.Range("A1:C1").Font.Bold = true;
                worksheet.Range("A1:C1").Interior.Color = Color.LightGray;

                worksheet.Range("E1:J1").Font.Bold = true;
                worksheet.Range("E1:J1").Interior.Color = Color.LightGray;

                worksheet.Range("E8:J8").Font.Bold = true;
                worksheet.Range("E8").Interior.Color = Color.LightBlue;

                // 设置数字格式
                worksheet.Range("C2:C6").NumberFormat = "¥#,##0";
                worksheet.Range("G2:G6").NumberFormat = "0";
                worksheet.Range("I2:I6").NumberFormat = "¥#,##0";
                worksheet.Range("J2:J6").NumberFormat = "¥#,##0";
                worksheet.Range("J8").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"LookupReferenceFunctionApplication_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示查找引用函数应用: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 查找引用函数应用时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 统计函数应用示例
        /// 演示STDEV、VAR、CORREL等统计函数的使用
        /// </summary>
        static void StatisticalFunctionApplicationExample()
        {
            Console.WriteLine("=== 统计函数应用示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "统计函数";

                // 创建销售数据
                worksheet.Range("A1").Value = "月份";
                worksheet.Range("B1").Value = "销售额";
                worksheet.Range("C1").Value = "广告费用";

                object[,] salesData = {
                    {"1月", 50000, 5000},
                    {"2月", 55000, 6000},
                    {"3月", 60000, 7000},
                    {"4月", 58000, 6500},
                    {"5月", 62000, 7200},
                    {"6月", 65000, 7800},
                    {"7月", 63000, 7500},
                    {"8月", 67000, 8000},
                    {"9月", 69000, 8200},
                    {"10月", 71000, 8500},
                    {"11月", 73000, 8800},
                    {"12月", 75000, 9000}
                };

                var dataRange = worksheet.Range("A2:C13");
                dataRange.Value = salesData;

                // 添加统计分析区域
                int statRow = 15;
                worksheet.Range($"A{statRow}").Value = "统计分析";
                worksheet.Range($"A{statRow}").Font.Bold = true;

                // 基本统计
                worksheet.Range($"A{statRow + 1}").Value = "销售额平均值";
                worksheet.Range($"B{statRow + 1}").Formula = "=AVERAGE(B2:B13)";

                worksheet.Range($"A{statRow + 2}").Value = "销售额标准差";
                worksheet.Range($"B{statRow + 2}").Formula = "=STDEV(B2:B13)";

                worksheet.Range($"A{statRow + 3}").Value = "销售额方差";
                worksheet.Range($"B{statRow + 3}").Formula = "=VAR(B2:B13)";

                worksheet.Range($"A{statRow + 4}").Value = "广告费用平均值";
                worksheet.Range($"B{statRow + 4}").Formula = "=AVERAGE(C2:C13)";

                worksheet.Range($"A{statRow + 5}").Value = "广告费用标准差";
                worksheet.Range($"B{statRow + 5}").Formula = "=STDEV(C2:C13)";

                // 相关性分析
                worksheet.Range($"A{statRow + 7}").Value = "相关性分析";
                worksheet.Range($"A{statRow + 7}").Font.Bold = true;

                worksheet.Range($"A{statRow + 8}").Value = "销售额与广告费用相关系数";
                worksheet.Range($"B{statRow + 8}").Formula = "=CORREL(B2:B13,C2:C13)";

                // 回归分析
                worksheet.Range($"A{statRow + 10}").Value = "回归分析";
                worksheet.Range($"A{statRow + 10}").Font.Bold = true;

                worksheet.Range($"A{statRow + 11}").Value = "斜率";
                worksheet.Range($"B{statRow + 11}").Formula = "=SLOPE(B2:B13,C2:C13)";

                worksheet.Range($"A{statRow + 12}").Value = "截距";
                worksheet.Range($"B{statRow + 12}").Formula = "=INTERCEPT(B2:B13,C2:C13)";

                worksheet.Range($"A{statRow + 13}").Value = "拟合度(R²)";
                worksheet.Range($"B{statRow + 13}").Formula = "=RSQ(B2:B13,C2:C13)";

                // 设置格式
                worksheet.Range("A1:C1").Font.Bold = true;
                worksheet.Range("A1:C1").Interior.Color = Color.LightGray;

                worksheet.Range($"A{statRow}:B{statRow + 13}").Font.Bold = true;
                worksheet.Range($"A{statRow}:A{statRow + 13}").Interior.Color = Color.LightBlue;

                // 设置数字格式
                worksheet.Range("B2:C13").NumberFormat = "¥#,##0";
                worksheet.Range($"B{statRow + 1}:B{statRow + 5}").NumberFormat = "¥#,##0";
                worksheet.Range($"B{statRow + 8}").NumberFormat = "0.000";
                worksheet.Range($"B{statRow + 11}:B{statRow + 13}").NumberFormat = "0.000";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"StatisticalFunctionApplication_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示统计函数应用: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 统计函数应用时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 数组公式应用示例
        /// 演示数组公式的使用
        /// </summary>
        static void ArrayFormulaApplicationExample()
        {
            Console.WriteLine("=== 数组公式应用示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "数组公式";

                // 创建产品销售数据
                worksheet.Range("A1").Value = "产品";
                worksheet.Range("B1").Value = "单价";
                worksheet.Range("C1").Value = "销量";
                worksheet.Range("D1").Value = "折扣";
                worksheet.Range("E1").Value = "实际收入";

                object[,] productData = {
                    {"产品A", 100, 50, 0.1},
                    {"产品B", 200, 30, 0.15},
                    {"产品C", 150, 40, 0.05},
                    {"产品D", 300, 20, 0.2},
                    {"产品E", 250, 35, 0.12}
                };

                var dataRange = worksheet.Range("A2:E6");
                dataRange.Value = productData;

                // 使用数组公式计算实际收入（单价*销量*(1-折扣)）
                worksheet.Range("E2").FormulaArray = "=B2:B6*C2:C6*(1-D2:D6)";

                // 添加总收入统计
                worksheet.Range("A8").Value = "总收入";
                worksheet.Range("A8").Font.Bold = true;
                worksheet.Range("E8").Formula = "=SUM(E2:E6)";
                worksheet.Range("E8").Font.Bold = true;

                // 设置格式
                worksheet.Range("A1:E1").Font.Bold = true;
                worksheet.Range("A1:E1").Interior.Color = Color.LightGray;

                worksheet.Range("A8:E8").Font.Bold = true;
                worksheet.Range("A8").Interior.Color = Color.LightBlue;

                // 设置数字格式
                worksheet.Range("B2:D6").NumberFormat = "0.00";
                worksheet.Range("E2:E6").NumberFormat = "¥#,##0";
                worksheet.Range("E8").NumberFormat = "¥#,##0";

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"ArrayFormulaApplication_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示数组公式应用: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 数组公式应用时出错: {ex.Message}");
            }

            Console.WriteLine();
        }

        /// <summary>
        /// 公式错误处理示例
        /// 演示如何处理公式中的错误
        /// </summary>
        static void FormulaErrorHandlingExample()
        {
            Console.WriteLine("=== 公式错误处理示例 ===");

            try
            {
                // 创建Excel应用程序实例
                using var excelApp = ExcelFactory.BlankWorkbook();

                // 获取活动工作簿和工作表
                var workbook = excelApp.ActiveWorkbook;
                var worksheet = workbook.ActiveSheetWrap;
                worksheet.Name = "公式错误处理";

                // 创建可能产生错误的公式
                worksheet.Range("A1").Value = "可能出现错误的公式";
                worksheet.Range("A1").Font.Bold = true;

                // 除零错误
                worksheet.Range("A2").Value = "除零错误:";
                worksheet.Range("B2").Formula = "=1/0";

                // 名称错误
                worksheet.Range("A3").Value = "名称错误:";
                worksheet.Range("B3").Formula = "=UNKNOWN_FUNCTION(1,2)";

                // 值错误
                worksheet.Range("A4").Value = "值错误:";
                worksheet.Range("B4").Formula = "=LEFT(123,1)";

                // 引用错误
                worksheet.Range("A5").Value = "引用错误:";
                worksheet.Range("B5").Formula = "=A1000";

                // 处理错误的公式
                worksheet.Range("A7").Value = "错误处理公式";
                worksheet.Range("A7").Font.Bold = true;

                // 使用IFERROR处理除零错误
                worksheet.Range("A8").Value = "处理除零错误:";
                worksheet.Range("B8").Formula = "=IFERROR(1/0,\"计算错误\")";

                // 使用IFERROR处理名称错误
                worksheet.Range("A9").Value = "处理名称错误:";
                worksheet.Range("B9").Formula = "=IFERROR(UNKNOWN_FUNCTION(1,2),\"函数不存在\")";

                // 使用IF处理特定错误
                worksheet.Range("A10").Value = "条件处理错误:";
                worksheet.Range("B10").Formula = "=IF(ISERROR(1/0),\"存在错误\",\"计算正常\")";

                // 使用ISERROR检查错误
                worksheet.Range("A11").Value = "错误检查:";
                worksheet.Range("B11").Formula = "=IF(ISERROR(B2),\"B2单元格有错误\",\"B2单元格正常\")";

                // 设置格式
                worksheet.Range("A1:E1").Font.Bold = true;
                worksheet.Range("A1:E1").Interior.Color = Color.LightGray;

                worksheet.Range("A7:E7").Font.Bold = true;
                worksheet.Range("A7").Interior.Color = Color.LightBlue;

                // 自动调整列宽
                worksheet.Columns.AutoFit();

                // 保存工作簿
                string fileName = $"FormulaErrorHandling_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                workbook.SaveAs(fileName);

                Console.WriteLine($"✓ 成功演示公式错误处理: {fileName}");

                // 显示错误信息
                Console.WriteLine("公式错误示例:");
                for (int i = 2; i <= 5; i++)
                {
                    var value = worksheet.Range($"B{i}").Value;
                    if (value is string errorText && errorText.StartsWith("#"))
                    {
                        Console.WriteLine($"  {worksheet.Range($"A{i}").Value} {errorText}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ 公式错误处理时出错: {ex.Message}");
            }

            Console.WriteLine();
        }
    }
}