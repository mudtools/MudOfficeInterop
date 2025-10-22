# 第六篇：公式与函数应用 - 深入Excel计算引擎

## 引言：Excel自动化的"智慧大脑"

如果说数据是Excel的"血液"，那么公式和函数就是Excel的"智慧大脑"！它们赋予了Excel强大的计算能力、逻辑判断能力和数据分析能力。在MudTools.OfficeInterop.Excel项目中，我们不仅保留了Excel原生的公式功能，还通过编程接口让这些功能变得更加强大和灵活。

想象一下这样的场景：你需要计算一个复杂项目的投资回报率，涉及多个变量、复杂的数学公式和条件判断。手动计算不仅容易出错，而且当某个变量发生变化时，需要重新计算所有相关数据。但通过公式自动化，你只需要设置一次公式，Excel就会自动完成所有的计算工作！

更令人兴奋的是，通过编程接口，你可以动态生成公式、批量设置公式、甚至创建自定义的复杂计算逻辑。这就像是给Excel装上了人工智能，让它能够理解你的业务逻辑，自动完成复杂的计算任务。

本篇将带你深入探索Excel公式和函数的奥秘，从基础的四则运算到高级的统计分析，从简单的条件判断到复杂的数组公式。准备好让你的Excel自动化系统拥有"智慧大脑"了吗？

## 基础公式操作

### 设置和读取公式

```csharp
public class FormulaManager
{
    public void BasicFormulaOperations(IExcelWorksheet worksheet)
    {
        // 设置基本公式
        worksheet["A1"].Value = "单价";
        worksheet["B1"].Value = "数量";
        worksheet["C1"].Value = "总价";
        
        // 设置数据
        worksheet["A2"].Value = 10.5;
        worksheet["B2"].Value = 5;
        
        // 设置公式
        worksheet["C2"].Formula = "=A2*B2";
        
        // 读取公式
        var formula = worksheet["C2"].Formula;
        var value = worksheet["C2"].Value;
        
        Console.WriteLine($"公式: {formula}");
        Console.WriteLine($"计算结果: {value}");
    }
    
    public void FormulaWithReferences(IExcelWorksheet worksheet)
    {
        // 使用相对引用
        for (int i = 1; i <= 5; i++)
        {
            worksheet[$"A{i}"].Value = i * 10;  // 单价
            worksheet[$"B{i}"].Value = i;       // 数量
            worksheet[$"C{i}"].Formula = $"=A{i}*B{i}";  // 总价
        }
    }
}
```

### 公式错误处理

```csharp
public class FormulaErrorHandler
{
    public void SafeFormulaOperations(IExcelWorksheet worksheet)
    {
        try
        {
            // 设置可能出错的公式
            worksheet["A1"].Formula = "=1/0";  // 除零错误
            
            // 检查公式错误
            var cell = worksheet["A1"];
            if (cell.HasFormula)
            {
                var formula = cell.Formula;
                var value = cell.Value;
                
                // 检查是否为错误值
                if (value is string errorText && errorText.StartsWith("#"))
                {
                    Console.WriteLine($"公式错误: {errorText}");
                    
                    // 处理特定错误
                    switch (errorText)
                    {
                        case "#DIV/0!":
                            Console.WriteLine("除零错误");
                            break;
                        case "#N/A":
                            Console.WriteLine("值不可用错误");
                            break;
                        case "#VALUE!":
                            Console.WriteLine("值错误");
                            break;
                        case "#REF!":
                            Console.WriteLine("引用错误");
                            break;
                        case "#NAME?":
                            Console.WriteLine("名称错误");
                            break;
                        case "#NUM!":
                            Console.WriteLine("数字错误");
                            break;
                        case "#NULL!":
                            Console.WriteLine("空值错误");
                            break;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"公式操作异常: {ex.Message}");
        }
    }
}
```

## 常用函数应用

### 数学和统计函数

```csharp
public class MathFunctionsManager
{
    public void ApplyMathFunctions(IExcelWorksheet worksheet)
    {
        // 准备测试数据
        var data = new double[] { 10, 20, 30, 40, 50, 60, 70, 80, 90, 100 };
        
        for (int i = 0; i < data.Length; i++)
        {
            worksheet[$"A{i + 1}"].Value = data[i];
        }
        
        // 应用统计函数
        worksheet["B1"].Value = "总和:";
        worksheet["B2"].Formula = "=SUM(A1:A10)";
        
        worksheet["C1"].Value = "平均值:";
        worksheet["C2"].Formula = "=AVERAGE(A1:A10)";
        
        worksheet["D1"].Value = "最大值:";
        worksheet["D2"].Formula = "=MAX(A1:A10)";
        
        worksheet["E1"].Value = "最小值:";
        worksheet["E2"].Formula = "=MIN(A1:A10)";
        
        worksheet["F1"].Value = "计数:";
        worksheet["F2"].Formula = "=COUNT(A1:A10)";
        
        worksheet["G1"].Value = "标准差:";
        worksheet["G2"].Formula = "=STDEV(A1:A10)";
        
        worksheet["H1"].Value = "方差:";
        worksheet["H2"].Formula = "=VAR(A1:A10)";
    }
    
    public void AdvancedMathFunctions(IExcelWorksheet worksheet)
    {
        // 三角函数
        worksheet["A1"].Value = "角度";
        worksheet["B1"].Value = "正弦";
        worksheet["C1"].Value = "余弦";
        worksheet["D1"].Value = "正切";
        
        for (int i = 0; i < 10; i++)
        {
            double angle = i * 10;
            worksheet[$"A{i + 2}"].Value = angle;
            worksheet[$"B{i + 2}"].Formula = $"=SIN(RADIANS(A{i + 2}))";
            worksheet[$"C{i + 2}"].Formula = $"=COS(RADIANS(A{i + 2}))";
            worksheet[$"D{i + 2}"].Formula = $"=TAN(RADIANS(A{i + 2}))";
        }
        
        // 对数函数
        worksheet["F1"].Value = "数值";
        worksheet["G1"].Value = "自然对数";
        worksheet["H1"].Value = "常用对数";
        
        for (int i = 0; i < 5; i++)
        {
            double value = Math.Pow(10, i);
            worksheet[$"F{i + 2}"].Value = value;
            worksheet[$"G{i + 2}"].Formula = $"=LN(F{i + 2})";
            worksheet[$"H{i + 2}"].Formula = $"=LOG10(F{i + 2})";
        }
    }
}
```

### 文本和日期函数

```csharp
public class TextDateFunctionsManager
{
    public void ApplyTextFunctions(IExcelWorksheet worksheet)
    {
        // 文本函数示例
        worksheet["A1"].Value = "原始文本";
        worksheet["B1"].Value = "大写";
        worksheet["C1"].Value = "小写";
        worksheet["D1"].Value = "长度";
        worksheet["E1"].Value = "左截取";
        worksheet["F1"].Value = "右截取";
        
        var testTexts = new string[] 
        { 
            "Hello World", 
            "Excel Automation", 
            "MudTools Library", 
            "C# Programming" 
        };
        
        for (int i = 0; i < testTexts.Length; i++)
        {
            worksheet[$"A{i + 2}"].Value = testTexts[i];
            worksheet[$"B{i + 2}"].Formula = $"=UPPER(A{i + 2})";
            worksheet[$"C{i + 2}"].Formula = $"=LOWER(A{i + 2})";
            worksheet[$"D{i + 2}"].Formula = $"=LEN(A{i + 2})";
            worksheet[$"E{i + 2}"].Formula = $"=LEFT(A{i + 2}, 5)";
            worksheet[$"F{i + 2}"].Formula = $"=RIGHT(A{i + 2}, 5)";
        }
        
        // 查找和替换函数
        worksheet["H1"].Value = "查找位置";
        worksheet["I1"].Value = "替换结果";
        
        worksheet["H2"].Formula = "=FIND("World", A2)";
        worksheet["I2"].Formula = "=SUBSTITUTE(A2, "Hello", "Hi")";
    }
    
    public void ApplyDateFunctions(IExcelWorksheet worksheet)
    {
        // 日期函数示例
        worksheet["A1"].Value = "当前日期";
        worksheet["B1"].Value = "当前时间";
        worksheet["C1"].Value = "年";
        worksheet["D1"].Value = "月";
        worksheet["E1"].Value = "日";
        worksheet["F1"].Value = "星期";
        
        worksheet["A2"].Formula = "=TODAY()";
        worksheet["B2"].Formula = "=NOW()";
        worksheet["C2"].Formula = "=YEAR(A2)";
        worksheet["D2"].Formula = "=MONTH(A2)";
        worksheet["E2"].Formula = "=DAY(A2)";
        worksheet["F2"].Formula = "=WEEKDAY(A2)";
        
        // 日期计算
        worksheet["A4"].Value = "开始日期";
        worksheet["B4"].Value = "结束日期";
        worksheet["C4"].Value = "天数差";
        worksheet["D4"].Value = "工作日";
        
        worksheet["A5"].Value = new DateTime(2024, 1, 1);
        worksheet["B5"].Value = new DateTime(2024, 12, 31);
        worksheet["C5"].Formula = "=B5-A5";
        worksheet["D5"].Formula = "=NETWORKDAYS(A5, B5)";
        
        // 日期格式化
        worksheet["E4"].Value = "格式化日期";
        worksheet["E5"].Formula = "=TEXT(A5, "yyyy年mm月dd日")";
    }
}
```

### 逻辑和查找函数

```csharp
public class LogicLookupFunctionsManager
{
    public void ApplyLogicFunctions(IExcelWorksheet worksheet)
    {
        // IF函数应用
        worksheet["A1"].Value = "分数";
        worksheet["B1"].Value = "等级";
        
        var scores = new int[] { 85, 92, 78, 65, 95, 58, 72, 88 };
        
        for (int i = 0; i < scores.Length; i++)
        {
            worksheet[$"A{i + 2}"].Value = scores[i];
            worksheet[$"B{i + 2}"].Formula = $"=IF(A{i + 2}>=90, "优秀", IF(A{i + 2}>=80, "良好", IF(A{i + 2}>=70, "中等", IF(A{i + 2}>=60, "及格", "不及格"))))";
        }
        
        // AND/OR函数
        worksheet["D1"].Value = "条件1";
        worksheet["E1"].Value = "条件2";
        worksheet["F1"].Value = "AND结果";
        worksheet["G1"].Value = "OR结果";
        
        worksheet["D2"].Value = true;
        worksheet["E2"].Value = false;
        worksheet["F2"].Formula = "=AND(D2, E2)";
        worksheet["G2"].Formula = "=OR(D2, E2)";
    }
    
    public void ApplyLookupFunctions(IExcelWorksheet worksheet)
    {
        // VLOOKUP函数示例
        // 创建查找表
        worksheet["H1"].Value = "产品ID";
        worksheet["I1"].Value = "产品名称";
        worksheet["J1"].Value = "价格";
        
        var products = new[]
        {
            new { ID = "P001", Name = "笔记本电脑", Price = 5999 },
            new { ID = "P002", Name = "智能手机", Price = 2999 },
            new { ID = "P003", Name = "平板电脑", Price = 1999 },
            new { ID = "P004", Name = "显示器", Price = 1299 }
        };
        
        for (int i = 0; i < products.Length; i++)
        {
            worksheet[$"H{i + 2}"].Value = products[i].ID;
            worksheet[$"I{i + 2}"].Value = products[i].Name;
            worksheet[$"J{i + 2}"].Value = products[i].Price;
        }
        
        // 使用VLOOKUP查找
        worksheet["L1"].Value = "查找ID";
        worksheet["M1"].Value = "产品名称";
        worksheet["N1"].Value = "价格";
        
        worksheet["L2"].Value = "P002";
        worksheet["M2"].Formula = "=VLOOKUP(L2, H2:J5, 2, FALSE)";
        worksheet["N2"].Formula = "=VLOOKUP(L2, H2:J5, 3, FALSE)";
        
        // HLOOKUP函数示例
        worksheet["P1"].Value = "月份";
        worksheet["P2"].Value = "一月";
        worksheet["P3"].Value = "二月";
        worksheet["P4"].Value = "三月";
        
        worksheet["Q1"].Value = "销售额";
        worksheet["Q2"].Value = 10000;
        worksheet["Q3"].Value = 12000;
        worksheet["Q4"].Value = 15000;
        
        worksheet["S1"].Value = "查找月份";
        worksheet["S2"].Value = "二月";
        worksheet["T1"].Value = "销售额";
        worksheet["T2"].Formula = "=HLOOKUP(S2, P1:Q4, 2, FALSE)";
    }
    
    public void ApplyIndexMatchFunctions(IExcelWorksheet worksheet)
    {
        // INDEX/MATCH组合 - 更强大的查找
        // 创建数据表
        worksheet["A10"].Value = "员工ID";
        worksheet["B10"].Value = "姓名";
        worksheet["C10"].Value = "部门";
        worksheet["D10"].Value = "工资";
        
        var employees = new[]
        {
            new { ID = "E001", Name = "张三", Department = "技术部", Salary = 8000 },
            new { ID = "E002", Name = "李四", Department = "销售部", Salary = 7000 },
            new { ID = "E003", Name = "王五", Department = "技术部", Salary = 8500 },
            new { ID = "E004", Name = "赵六", Department = "人事部", Salary = 7500 }
        };
        
        for (int i = 0; i < employees.Length; i++)
        {
            worksheet[$"A{i + 11}"].Value = employees[i].ID;
            worksheet[$"B{i + 11}"].Value = employees[i].Name;
            worksheet[$"C{i + 11}"].Value = employees[i].Department;
            worksheet[$"D{i + 11}"].Value = employees[i].Salary;
        }
        
        // 使用INDEX/MATCH查找
        worksheet["F10"].Value = "查找姓名";
        worksheet["G10"].Value = "李四";
        worksheet["F11"].Value = "部门";
        worksheet["F12"].Value = "工资";
        
        // 查找部门
        worksheet["G11"].Formula = "=INDEX(C11:C14, MATCH(G10, B11:B14, 0))";
        // 查找工资
        worksheet["G12"].Formula = "=INDEX(D11:D14, MATCH(G10, B11:B14, 0))";
    }
}
```

## 数组公式应用

### 单单元格数组公式

```csharp
public class ArrayFormulaManager
{
    public void SingleCellArrayFormulas(IExcelWorksheet worksheet)
    {
        // 单单元格数组公式 - 计算总销售额
        worksheet["A1"].Value = "产品";
        worksheet["B1"].Value = "单价";
        worksheet["C1"].Value = "数量";
        worksheet["D1"].Value = "销售额";
        
        var products = new[]
        {
            new { Name = "产品A", Price = 100, Quantity = 50 },
            new { Name = "产品B", Price = 200, Quantity = 30 },
            new { Name = "产品C", Price = 150, Quantity = 40 }
        };
        
        for (int i = 0; i < products.Length; i++)
        {
            worksheet[$"A{i + 2}"].Value = products[i].Name;
            worksheet[$"B{i + 2}"].Value = products[i].Price;
            worksheet[$"C{i + 2}"].Value = products[i].Quantity;
            worksheet[$"D{i + 2}"].Formula = $"=B{i + 2}*C{i + 2}";
        }
        
        // 使用数组公式计算总销售额
        worksheet["F1"].Value = "总销售额";
        worksheet["F2"].FormulaArray = "=SUM(B2:B4*C2:C4)";
        
        // 使用数组公式计算加权平均单价
        worksheet["G1"].Value = "加权平均单价";
        worksheet["G2"].FormulaArray = "=SUM(B2:B4*C2:C4)/SUM(C2:C4)";
    }
    
    public void MultiCellArrayFormulas(IExcelWorksheet worksheet)
    {
        // 多单元格数组公式 - 批量计算
        var data = new double[] { 10, 20, 30, 40, 50 };
        
        // 输入数据
        for (int i = 0; i < data.Length; i++)
        {
            worksheet[$"A{i + 1}"].Value = data[i];
        }
        
        // 使用数组公式批量计算平方和立方
        var outputRange = worksheet.Range("B1:C5");
        
        // 注意：这种方式设置的是普通公式，不是真正的数组公式
        // 真正的数组公式需要在Excel中手动设置Ctrl+Shift+Enter
        for (int i = 0; i < data.Length; i++)
        {
            worksheet[$"B{i + 1}"].Formula = $"=A{i + 1}^2";  // 平方
            worksheet[$"C{i + 1}"].Formula = $"=A{i + 1}^3";  // 立方
        }
        
        // 真正的数组公式示例（需要在Excel中确认）
        worksheet["E1"].Value = "数组公式结果";
        // worksheet.Range("E2:E6").FormulaArray = "=A1:A5*2"; // 这行代码在实际中可能需要特殊处理
    }
}
```

### 动态数组公式（Excel 365）

```csharp
public class DynamicArrayManager
{
    public void DynamicArrayOperations(IExcelWorksheet worksheet)
    {
        // 动态数组公式 - 适用于Excel 365
        // 创建源数据
        var sourceData = new[] { "苹果", "香蕉", "橙子", "葡萄", "西瓜" };
        
        for (int i = 0; i < sourceData.Length; i++)
        {
            worksheet[$"A{i + 1}"].Value = sourceData[i];
        }
        
        // 使用动态数组函数（需要Excel 365支持）
        worksheet["C1"].Value = "排序结果";
        // worksheet["C2"].Formula = "=SORT(A1:A5)"; // 动态排序
        
        worksheet["D1"].Value = "唯一值";
        // worksheet["D2"].Formula = "=UNIQUE(A1:A5)"; // 提取唯一值
        
        // 实际应用中可能需要版本检测
        TrySetDynamicFormula(worksheet, "C2", "=SORT(A1:A5)");
        TrySetDynamicFormula(worksheet, "D2", "=UNIQUE(A1:A5)");
    }
    
    private void TrySetDynamicFormula(IExcelWorksheet worksheet, string cellAddress, string formula)
    {
        try
        {
            worksheet[cellAddress].Formula = formula;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"动态公式设置失败: {ex.Message}");
            // 回退到传统方法
            worksheet[cellAddress].Value = "动态公式不支持";
        }
    }
}
```

## 自定义函数和高级应用

### 通过VBA调用自定义函数

```csharp
public class CustomFunctionManager
{
    public void CallVbaCustomFunctions(IExcelWorksheet worksheet)
    {
        // 假设工作簿中已经存在自定义VBA函数
        // 例如：一个计算税金的函数 CalculateTax(amount, rate)
        
        try
        {
            worksheet["A1"].Value = "金额";
            worksheet["B1"].Value = "税率";
            worksheet["C1"].Value = "税金";
            
            worksheet["A2"].Value = 1000;
            worksheet["B2"].Value = 0.1; // 10%
            
            // 调用自定义VBA函数
            worksheet["C2"].Formula = "=CalculateTax(A2, B2)";
            
            Console.WriteLine("自定义函数调用成功");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"自定义函数调用失败: {ex.Message}");
            // 使用内置函数替代
            worksheet["C2"].Formula = "=A2*B2";
        }
    }
    
    public void ComplexFormulaBuilder(IExcelWorksheet worksheet)
    {
        // 构建复杂公式
        var formulaBuilder = new StringBuilder();
        
        // 嵌套IF函数示例
        formulaBuilder.Append("=IF(A1>90,"优秀",");
        formulaBuilder.Append("IF(A1>80,"良好",");
        formulaBuilder.Append("IF(A1>70,"中等",");
        formulaBuilder.Append("IF(A1>60,"及格","不及格")))");
        
        worksheet["B1"].Formula = formulaBuilder.ToString();
        worksheet["A1"].Value = 85;
        
        // 复杂数学公式
        worksheet["C1"].Value = "复杂计算";
        worksheet["C2"].Formula = "=SUM(A1:A10)*AVERAGE(B1:B10)/COUNT(C1:C10)";
    }
}
```

### 公式性能优化

```csharp
public class FormulaPerformanceOptimizer
{
    public void OptimizeFormulaPerformance(IExcelWorksheet worksheet)
    {
        // 1. 避免易失性函数的过度使用
        // 易失性函数：NOW(), TODAY(), RAND(), OFFSET(), INDIRECT()等
        
        // 不好的做法：在大量单元格中使用易失性函数
        for (int i = 1; i <= 1000; i++)
        {
            // worksheet[$"A{i}"].Formula = "=NOW()"; // 每次计算都会更新
        }
        
        // 好的做法：在一个单元格中使用，然后引用
        worksheet["B1"].Formula = "=NOW()";
        for (int i = 1; i <= 1000; i++)
        {
            worksheet[$"C{i}"].Formula = "=B1"; // 只计算一次
        }
        
        // 2. 使用辅助列减少复杂计算
        SetupHelperColumns(worksheet);
        
        // 3. 避免循环引用
        AvoidCircularReferences(worksheet);
    }
    
    private void SetupHelperColumns(IExcelWorksheet worksheet)
    {
        // 使用辅助列优化复杂公式
        var salesData = new[]
        {
            new { Product = "产品A", Q1 = 100, Q2 = 120, Q3 = 110, Q4 = 130 },
            new { Product = "产品B", Q1 = 200, Q2 = 180, Q3 = 220, Q4 = 210 },
            new { Product = "产品C", Q1 = 150, Q2 = 160, Q3 = 170, Q4 = 180 }
        };
        
        // 设置数据
        for (int i = 0; i < salesData.Length; i++)
        {
            worksheet[$"A{i + 1}"].Value = salesData[i].Product;
            worksheet[$"B{i + 1}"].Value = salesData[i].Q1;
            worksheet[$"C{i + 1}"].Value = salesData[i].Q2;
            worksheet[$"D{i + 1}"].Value = salesData[i].Q3;
            worksheet[$"E{i + 1}"].Value = salesData[i].Q4;
        }
        
        // 使用辅助列计算季度平均（而不是在最终公式中重复计算）
        for (int i = 0; i < salesData.Length; i++)
        {
            worksheet[$"F{i + 1}"].Formula = $"=AVERAGE(B{i + 1}:E{i + 1})"; // 辅助列
        }
        
        // 最终计算使用辅助列
        worksheet["G1"].Value = "年度表现";
        for (int i = 0; i < salesData.Length; i++)
        {
            worksheet[$"G{i + 1}"].Formula = $"=IF(F{i + 1}>150, "优秀", IF(F{i + 1}>100, "良好", "需改进"))";
        }
    }
    
    private void AvoidCircularReferences(IExcelWorksheet worksheet)
    {
        // 避免循环引用示例
        // 错误的循环引用
        // worksheet["A1"].Formula = "=B1+1";
        // worksheet["B1"].Formula = "=A1+1"; // 循环引用！
        
        // 正确的做法
        worksheet["A2"].Value = 10;
        worksheet["B2"].Formula = "=A2+1";
        worksheet["C2"].Formula = "=B2+1"; // 线性引用，无循环
    }
}
```

## 实际应用场景

### 财务报表自动化

```csharp
public class FinancialReportManager
{
    public void CreateFinancialReport(IExcelWorksheet worksheet)
    {
        // 创建财务报表模板
        SetupFinancialTemplate(worksheet);
        
        // 设置财务公式
        SetupFinancialFormulas(worksheet);
        
        // 数据验证和保护
        ProtectFinancialWorksheet(worksheet);
    }
    
    private void SetupFinancialTemplate(IExcelWorksheet worksheet)
    {
        // 设置财务报表结构
        worksheet["A1"].Value = "月度财务报表";
        worksheet["A1"].Font.Bold = true;
        worksheet["A1"].Font.Size = 16;
        
        // 收入部分
        worksheet["A3"].Value = "收入项目";
        worksheet["B3"].Value = "金额";
        
        var revenueItems = new[]
        {
            "产品销售收入", "服务收入", "其他收入"
        };
        
        for (int i = 0; i < revenueItems.Length; i++)
        {
            worksheet[$"A{i + 4}"].Value = revenueItems[i];
            worksheet[$"B{i + 4}"].Value = 0; // 初始值
        }
        
        worksheet["A7"].Value = "总收入";
        worksheet["A7"].Font.Bold = true;
        worksheet["B7"].Formula = "=SUM(B4:B6)";
        
        // 支出部分
        worksheet["D3"].Value = "支出项目";
        worksheet["E3"].Value = "金额";
        
        var expenseItems = new[]
        {
            "人员成本", "物料成本", "运营费用", "其他支出"
        };
        
        for (int i = 0; i < expenseItems.Length; i++)
        {
            worksheet[$"D{i + 4}"].Value = expenseItems[i];
            worksheet[$"E{i + 4}"].Value = 0; // 初始值
        }
        
        worksheet["D8"].Value = "总支出";
        worksheet["D8"].Font.Bold = true;
        worksheet["E8"].Formula = "=SUM(E4:E7)";
        
        // 利润计算
        worksheet["A10"].Value = "净利润";
        worksheet["A10"].Font.Bold = true;
        worksheet["B10"].Font.Bold = true;
        worksheet["B10"].Formula = "=B7-E8";
        
        // 利润率
        worksheet["A11"].Value = "利润率";
        worksheet["B11"].Formula = "=IF(B7>0, B10/B7, 0)";
        worksheet["B11"].NumberFormat = "0.00%";
    }
    
    private void SetupFinancialFormulas(IExcelWorksheet worksheet)
    {
        // 设置复杂的财务公式
        
        // 月度增长率计算
        worksheet["D10"].Value = "上月收入";
        worksheet["E10"].Value = 0;
        worksheet["D11"].Value = "收入增长率";
        worksheet["E11"].Formula = "=IF(E10>0, (B7-E10)/E10, 0)";
        worksheet["E11"].NumberFormat = "0.00%";
        
        // 预算完成率
        worksheet["D12"].Value = "收入预算";
        worksheet["E12"].Value = 0;
        worksheet["D13"].Value = "预算完成率";
        worksheet["E13"].Formula = "=IF(E12>0, B7/E12, 0)";
        worksheet["E13"].NumberFormat = "0.00%";
        
        // 财务比率分析
        SetupFinancialRatios(worksheet);
    }
    
    private void SetupFinancialRatios(IExcelWorksheet worksheet)
    {
        // 财务比率分析
        worksheet["G3"].Value = "财务比率分析";
        worksheet["G3"].Font.Bold = true;
        
        var ratios = new[]
        {
            new { Name = "流动比率", Formula = "=B7/E8" },
            new { Name = "毛利率", Formula = "=IF(B7>0, (B7-SUM(E4:E5))/B7, 0)" },
            new { Name = "净利率", Formula = "=B11" },
            new { Name = "支出收入比", Formula = "=IF(B7>0, E8/B7, 0)" }
        };
        
        for (int i = 0; i < ratios.Length; i++)
        {
            worksheet[$"G{i + 4}"].Value = ratios[i].Name;
            worksheet[$"H{i + 4}"].Formula = ratios[i].Formula;
            worksheet[$"H{i + 4}"].NumberFormat = "0.00";
        }
    }
    
    private void ProtectFinancialWorksheet(IExcelWorksheet worksheet)
    {
        // 保护公式单元格
        var formulaCells = new[] { "B7", "B10", "B11", "E8", "E11", "E13" };
        
        foreach (var cellAddress in formulaCells)
        {
            var cell = worksheet[cellAddress];
            cell.Locked = true; // 锁定公式单元格
        }
        
        // 设置数据输入区域为可编辑
        var inputRanges = new[] { "B4:B6", "E4:E7", "E10", "E12" };
        
        foreach (var rangeAddress in inputRanges)
        {
            var range = worksheet.Range(rangeAddress);
            range.Locked = false; // 解锁输入区域
        }
        
        // 应用工作表保护
        worksheet.Protect("financial123");
    }
}
```

### 销售数据分析

```csharp
public class SalesAnalysisManager
{
    public void CreateSalesAnalysis(IExcelWorksheet worksheet)
    {
        // 创建销售数据分析报表
        SetupSalesData(worksheet);
        
        // 应用分析公式
        ApplyAnalysisFormulas(worksheet);
        
        // 创建动态分析区域
        CreateDynamicAnalysis(worksheet);
    }
    
    private void SetupSalesData(IExcelWorksheet worksheet)
    {
        // 设置销售数据
        var salesData = new[]
        {
            new { Region = "华北", Product = "产品A", Month = "1月", Sales = 1000 },
            new { Region = "华北", Product = "产品B", Month = "1月", Sales = 1500 },
            new { Region = "华东", Product = "产品A", Month = "1月", Sales = 1200 },
            new { Region = "华东", Product = "产品B", Month = "1月", Sales = 1800 },
            new { Region = "华南", Product = "产品A", Month = "1月", Sales = 900 },
            new { Region = "华南", Product = "产品B", Month = "1月", Sales = 1300 }
        };
        
        // 表头
        worksheet["A1"].Value = "区域";
        worksheet["B1"].Value = "产品";
        worksheet["C1"].Value = "月份";
        worksheet["D1"].Value = "销售额";
        
        // 数据行
        for (int i = 0; i < salesData.Length; i++)
        {
            worksheet[$"A{i + 2}"].Value = salesData[i].Region;
            worksheet[$"B{i + 2}"].Value = salesData[i].Product;
            worksheet[$"C{i + 2}"].Value = salesData[i].Month;
            worksheet[$"D{i + 2}"].Value = salesData[i].Sales;
        }
    }
    
    private void ApplyAnalysisFormulas(IExcelWorksheet worksheet)
    {
        // 区域分析
        worksheet["F1"].Value = "区域销售汇总";
        worksheet["F1"].Font.Bold = true;
        
        var regions = new[] { "华北", "华东", "华南" };
        
        for (int i = 0; i < regions.Length; i++)
        {
            worksheet[$"F{i + 2}"].Value = regions[i];
            worksheet[$"G{i + 2}"].Formula = $"=SUMIF(A2:A7, F{i + 2}, D2:D7)";
        }
        
        // 产品分析
        worksheet["I1"].Value = "产品销售汇总";
        worksheet["I1"].Font.Bold = true;
        
        var products = new[] { "产品A", "产品B" };
        
        for (int i = 0; i < products.Length; i++)
        {
            worksheet[$"I{i + 2}"].Value = products[i];
            worksheet[$"J{i + 2}"].Formula = $"=SUMIF(B2:B7, I{i + 2}, D2:D7)";
        }
        
        // 排名分析
        worksheet["L1"].Value = "销售排名";
        worksheet["L1"].Font.Bold = true;
        
        for (int i = 0; i < 6; i++)
        {
            worksheet[$"L{i + 2}"].Value = $"第{i + 1}名";
            worksheet[$"M{i + 2}"].Formula = $"=LARGE(D2:D7, {i + 1})";
        }
    }
    
    private void CreateDynamicAnalysis(IExcelWorksheet worksheet)
    {
        // 创建动态分析区域
        worksheet["O1"].Value = "动态分析";
        worksheet["O1"].Font.Bold = true;
        
        // 平均值分析
        worksheet["O2"].Value = "平均销售额";
        worksheet["P2"].Formula = "=AVERAGE(D2:D7)";
        
        // 标准差分析
        worksheet["O3"].Value = "销售标准差";
        worksheet["P3"].Formula = "=STDEV(D2:D7)";
        
        // 变异系数
        worksheet["O4"].Value = "变异系数";
        worksheet["P4"].Formula = "=IF(P2>0, P3/P2, 0)";
        
        // 增长预测
        worksheet["O5"].Value = "下月预测";
        worksheet["P5"].Formula = "=P2*1.1"; // 假设10%增长
    }
}
```

## 总结

本文详细介绍了MudTools.OfficeInterop.Excel项目中公式和函数的应用，涵盖了从基础公式操作到高级函数应用的完整内容。通过实际的代码示例，展示了如何：

1. **基础公式操作** - 设置、读取和处理公式
2. **常用函数应用** - 数学、统计、文本、日期、逻辑和查找函数
3. **数组公式** - 单单元格和多单元格数组公式的应用
4. **自定义函数** - 通过VBA调用自定义函数
5. **性能优化** - 公式性能优化技巧
6. **实际应用** - 财务报表和销售数据分析的实际案例

这些功能为开发者提供了强大的工具，可以创建复杂的计算和分析应用，满足企业级Excel自动化的各种需求。