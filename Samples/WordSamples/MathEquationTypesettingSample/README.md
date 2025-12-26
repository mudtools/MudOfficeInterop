# 数学公式排版示例项目

这个示例项目展示了如何使用 MudTools.OfficeInterop.Word 进行各种数学公式的创建、格式化和自动化排版。

## 项目结构

```
MathEquationTypesettingSample/
├── Program.cs                    # 主程序，包含所有示例
├── LaTeXToWordConverter.cs       # LaTeX到Word公式转换器
├── ScientificPaperFormatter.cs   # 学术论文格式化器
├── README.md                     # 项目说明文档
└── MathEquationTypesettingSample.csproj # 项目配置文件
```

## 功能特性

### 1. 基础公式创建
- 创建简单的数学公式
- 设置公式显示格式（专业格式 vs 内联格式）
- 基本的公式构建操作

### 2. 分数和根式
- 创建各种类型的分数（横线、斜线、无分数线）
- 平方根和n次根式的创建
- 根式参数设置

### 3. 积分和求和
- 定积分和不定积分的创建
- 求和符号和乘积符号
- 上下标的设置

### 4. 矩阵排版
- 各种矩阵格式的创建
- 行列间距和对齐设置
- 复杂矩阵元素的填充

### 5. 方程组
- 多行方程组的排版
- 对齐方式的控制
- 方程组编号管理

### 6. 嵌套公式
- 复杂嵌套结构的创建
- 分数内嵌套根式
- 多层上标下标

### 7. 样式和格式控制
- 字体、大小、颜色设置
- 对齐和布局控制
- 段落格式管理

### 8. LaTeX转Word公式
- LaTeX语法解析
- 常用LaTeX命令转换
- 自动构建Word公式结构

### 9. 学术论文自动化排版
- 批量公式处理
- 自动编号和书签
- 期刊样式应用

## 运行示例

### 环境要求
- .NET Framework 4.6.2 或更高版本
- Microsoft Office Word（支持COM互操作）
- MudTools.OfficeInterop.Word NuGet包

### 编译和运行
```bash
cd MathEquationTypesettingSample
dotnet restore
dotnet run
```

### 示例输出
运行程序后，将在当前目录生成以下示例文档：

1. **BasicEquationSample.docx** - 基础公式示例
2. **FractionAndRadicalSample.docx** - 分数和根式示例
3. **IntegralAndSumSample.docx** - 积分和求和示例
4. **MatrixTypesettingSample.docx** - 矩阵排版示例
5. **EquationSystemSample.docx** - 方程组示例
6. **NestedEquationSample.docx** - 嵌套公式示例
7. **EquationFormattingSample.docx** - 格式控制示例
8. **LaTeXToWordSample.docx** - LaTeX转换示例
9. **ScientificPaperSample.docx** - 学术论文示例

## 代码示例

### 创建基础公式
```csharp
using var application = new WordApplication();
IWordDocument document = application.Documents.Add();
IWordRange range = document.Content;

IWordOMaths oMaths = range.OMaths;
IWordRange formulaRange = oMaths.Add(range);
IWordOMath oMath = oMaths[1];  // COM集合索引从1开始

oMath.Range.Text = "x^2 + y^2 = z^2";
oMath.BuildUp();
oMath.Type = WdOMathType.wdOMathDisplay;
```

### 创建分数
```csharp
var fractionFunction = oMath.Functions.Add(range, WdOMathFunctionType.wdOMathFunctionFrac);
var fraction = fractionFunction.Frac;

fraction.Num.Range.Text = "a^2 + b^2";
fraction.Den.Range.Text = "c^2";
fraction.Type = WdOMathFracType.wdOMathFracBar;
```

### 创建矩阵
```csharp
var matrixFunction = oMath.Functions.Add(range, WdOMathFunctionType.wdOMathFunctionMat);
var matrix = matrixFunction.Mat;

// 添加3x3矩阵
for (int i = 0; i < 3; i++)
{
    matrix.Rows.Add(null);
    matrix.Cols.Add(null);
}

// 填充矩阵元素
matrix.Cell(1, 1).Range.Text = "a";
matrix.Cell(1, 2).Range.Text = "b";
matrix.Cell(1, 3).Range.Text = "c";
// ... 其他元素
```

### LaTeX转Word公式
```csharp
var converter = new LaTeXToWordConverter();
IWordOMath oMath = converter.ConvertLaTeXToWordFormula(range, @"\frac{a^2 + b^2}{c^2}");
oMath.BuildUp();
```

### 学术论文格式化
```csharp
var formatter = new ScientificPaperFormatter();
var equations = new List<string>
{
    @"\frac{d^2y}{dx^2} + \omega^2 y = 0",
    @"\int_{0}^{\infty} e^{-x^2} dx = \frac{\sqrt{\pi}}{2}",
    @"\sum_{i=1}^{n} i^2 = \frac{n(n+1)(2n+1)}{6}"
};

formatter.FormatScientificPaper("template.docx", equations, "output.docx");
```

## 支持的LaTeX命令

当前版本的转换器支持以下LaTeX命令：

- `\frac{分子}{分母}` - 分数
- `\sqrt[根次]{表达式}` - 根式
- `\int_{下标}^{上标} 表达式` - 积分
- `\sum_{下标}^{上标} 表达式` - 求和
- `\begin{pmatrix} ... \end{pmatrix}` - 矩阵
- `x^{上标}` - 上标
- `x_{下标}` - 下标

## 期刊样式支持

项目预定义了多种期刊的格式要求：

- **IEEE** - 10pt Times New Roman，较小间距
- **Nature** - 12pt Times New Roman，标准间距
- **Science** - 11pt Times New Roman，中等间距
- **Default** - 12pt Times New Roman，较大间距

## 注意事项

1. **COM互操作要求**：确保系统上安装了Microsoft Office Word
2. **权限要求**：程序需要对当前目录的写入权限
3. **内存管理**：使用`using`语句确保COM对象正确释放
4. **错误处理**：示例中包含了基本的错误处理，生产环境可能需要更完善的异常处理

## 扩展功能

这个示例项目可以作为基础，进一步扩展以支持：

- 更多LaTeX命令的转换
- 公式渲染性能优化
- 自定义期刊样式模板
- 公式交叉引用管理
- 批量文档处理
- Web API集成

## 相关链接

- [MudTools.OfficeInterop.Word 文档](../../../README.md)
- [Word COM对象模型参考](https://docs.microsoft.com/en-us/office/vba/api/overview/word)
- [LaTeX数学模式文档](https://www.overleaf.com/learn/latex/Mathematical_expressions)

## 许可证

本项目遵循 MIT 许可证和 Apache 许可证（版本 2.0）。