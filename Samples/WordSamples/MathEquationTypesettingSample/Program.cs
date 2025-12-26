//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace MathEquationTypesettingSample
{
    /// <summary>
    /// 数学公式排版示例程序
    /// 演示如何使用 MudTools.OfficeInterop.Word 进行各种数学公式的创建和格式化
    /// </summary>
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("=== C# Word数学公式排版示例 ===\n");

            try
            {
                // 积分和求和示例
                await IntegralAndSumSample();

                // 分数和根式示例
                await FractionAndRadicalSample();

                // 基础公式创建示例
                await BasicEquationCreationSample();


                // 矩阵排版示例
                await MatrixTypesettingSample();

                // 方程组示例
                await EquationSystemSample();

                // 嵌套公式示例
                await NestedEquationSample();

                // 公式样式和格式控制示例
                await EquationFormattingSample();

                // LaTeX转Word公式示例
                await LaTeXToWordSample();

                // 学术论文自动化排版示例
                await ScientificPaperSample();

                Console.WriteLine("\n所有示例执行完成！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行过程中发生错误: {ex.Message}");
                Console.WriteLine($"详细信息: {ex}");
            }

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 基础公式创建示例
        /// </summary>
        static async Task BasicEquationCreationSample()
        {
            Console.WriteLine("1. 基础公式创建示例");

            using var application = WordFactory.BlankDocument();
            application.Visible = true;

            IWordDocument document = application.Documents.Add();

            IWordRange range = document.Content;

            // 插入数学公式
            IWordOMaths oMaths = range.OMaths;
            IWordRange formulaRange = oMaths.Add(range);

            // 设置公式内容
            IWordOMath oMath = oMaths[1];  // COM集合索引从1开始
            oMath.Range.Text = "x^2 + y^2 = z^2";

            // 构建专业格式显示
            oMath.Type = WdOMathType.wdOMathDisplay;
            oMath.BuildUp();
            //

            // 添加标题
            IWordRange titleRange = document.Range(0, 0);
            titleRange.Text = "基础数学公式示例：毕达哥拉斯定理\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;

            document.SaveAs2(Path.Combine(Environment.CurrentDirectory, "BasicEquationSample.docx"));
            document.Close();
            application.Quit();
            Console.WriteLine("   ✓ 基础公式创建完成，文件保存为: BasicEquationSample.docx");
        }

        /// <summary>
        /// 分数和根式示例
        /// </summary>
        static async Task FractionAndRadicalSample()
        {
            Console.WriteLine("2. 分数和根式示例");

            using var application = WordFactory.BlankDocument();
            application.Visible = false;

            IWordDocument document = application.Documents.Add();


            // 添加标题
            IWordRange titleRange = document.Content;
            titleRange.Text = "分数和根式示例\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;

            // 创建分数示例
            var range1 = WordHelper.GetEndRange(document);
            range1.InsertParagraphAfter();
            //range1.Collapse(WdCollapseDirection.wdCollapseEnd);
            range1.Text = "c^2/a^2 + b^2";

            IWordOMaths oMaths1 = range1.OMaths;
            IWordRange fractionRange = oMaths1.Add(range1);
            IWordOMath fractionOMath = oMaths1[1];  // COM集合索引从1开始
            if (fractionOMath.Functions.Count > 0)
            {
                var fracFunc = fractionOMath.Functions[1]; // 自动识别的分数函数
                fracFunc.Frac.Num.Range.Text = "a^2 + b^2";
                fracFunc.Frac.Den.Range.Text = "c^2";
                var fraction = fracFunc.Frac;
                fraction.Type = WdOMathFracType.wdOMathFracBar;
            }
            fractionOMath.BuildUp();
            document.OMaths.BuildUp();
            // 创建根式示例
            var range2 = WordHelper.GetEndRange(document);
            range2.InsertParagraphAfter();
            range2.Collapse(WdCollapseDirection.wdCollapseEnd);
            range2.Text = @"\sqrt( )";
            // 创建根式函数
            IWordOMaths oMaths2 = range2.OMaths;
            IWordRange radicalRange = oMaths2.Add(range2);
            IWordOMath radicalOMath = oMaths2[oMaths2.Count];  // 获取刚添加的公式
            radicalOMath.Functions[1].Args[1].Range.Text = "x^2 + 2xy + y^2";


            // 构建专业格式
            //fractionOMath.BuildUp();
            radicalOMath.BuildUp();

            document.SaveAs2(Path.Combine(Environment.CurrentDirectory, "FractionAndRadicalSample.docx"));
            document.Close();
            application.Quit();
            Console.WriteLine("   ✓ 分数和根式示例完成，文件保存为: FractionAndRadicalSample.docx");
        }

        /// <summary>
        /// 积分和求和示例
        /// </summary>
        static async Task IntegralAndSumSample()
        {
            Console.WriteLine("3. 积分和求和示例");

            using var application = WordFactory.BlankDocument();
            application.Visible = false;

            IWordDocument document = application.Documents.Add();

            // 添加标题
            IWordRange titleRange = document.Content;
            titleRange.Text = "积分和求和示例\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;

            // 创建积分示例
            var range1 = document.Range(document.Content.End - 1, document.Content.End);
            range1.InsertParagraphAfter();
            range1.Collapse(WdCollapseDirection.wdCollapseEnd);
            range1.Text = @"(√(t_c1*t_w1)*2)/t_f1 ∙ (√(t_c2*t_w2)*2)/t_f2"; ;

            IWordOMaths oMaths1 = document.OMaths;
            IWordRange integralRange = oMaths1.Add(range1);
            IWordOMath integralOMath = oMaths1[1];  // COM集合索引从1开始

            //// 创建积分函数
            var integralFunction = integralOMath.Functions[1];
            var nary = integralFunction.Nary;

            //// 设置积分表达式
            nary.E.Range.Text = "e^{-x^2} dx";
            nary.Sub.Range.Text = "-∞";
            nary.Sup.Range.Text = "∞";
            nary.Char = '∫';

            //// 创建求和示例
            //var range2 = document.Range(document.Content.End - 1, document.Content.End);
            //range2.InsertParagraphAfter();
            //range2.Collapse(WdCollapseDirection.wdCollapseEnd);

            //IWordOMaths oMaths2 = range2.OMaths;
            //IWordRange sumRange = oMaths2.Add(range2);
            //IWordOMath sumOMath = oMaths2[1];  // COM集合索引从1开始

            //// 创建求和函数
            //var sumFunction = sumOMath.Functions.Add(sumRange, WdOMathFunctionType.wdOMathFunctionNary);
            //var sum = sumFunction.Nary;

            //// 设置求和表达式
            //sum.E.Range.Text = "i^2";
            //sum.Sub.Range.Text = "i=1";
            //sum.Sup.Range.Text = "n";
            //sum.Char = '∑';

            // 构建专业格式
            integralOMath.BuildUp();
            //sumOMath.BuildUp();

            document.SaveAs2(Path.Combine(Environment.CurrentDirectory, "IntegralAndSumSample.docx"));
            document.Close();

            Console.WriteLine("   ✓ 积分和求和示例完成，文件保存为: IntegralAndSumSample.docx");
        }

        /// <summary>
        /// 矩阵排版示例
        /// </summary>
        static async Task MatrixTypesettingSample()
        {
            Console.WriteLine("4. 矩阵排版示例");

            using var application = WordFactory.BlankDocument();
            application.Visible = false;

            IWordDocument document = application.Documents.Add();

            // 添加标题
            IWordRange titleRange = document.Content;
            titleRange.Text = "矩阵排版示例\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;

            // 创建矩阵
            var range = document.Range(document.Content.End - 1, document.Content.End);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            IWordOMaths oMaths = range.OMaths;
            IWordRange matrixRange = oMaths.Add(range);
            IWordOMath matrixOMath = oMaths[1];  // COM集合索引从1开始

            // 创建矩阵函数
            var matrixFunction = matrixOMath.Functions.Add(matrixRange, WdOMathFunctionType.wdOMathFunctionMat);
            var matrix = matrixFunction.Mat;

            // 添加行和列
            for (int i = 0; i < 3; i++)
            {
                matrix.Rows.Add(null);
            }
            for (int j = 0; j < 3; j++)
            {
                matrix.Cols.Add(null);
            }

            // 设置矩阵元素
            string[,] matrixElements = {
                { "a", "b", "c" },
                { "d", "e", "f" },
                { "g", "h", "i" }
            };

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    matrix.Cell(row + 1, col + 1).Range.Text = matrixElements[row, col];
                }
            }

            // 设置矩阵格式
            matrix.Align = WdOMathVertAlignType.wdOMathVertAlignCenter;
            matrix.RowSpacing = 20;
            matrix.ColSpacing = 15;

            // 构建专业格式
            matrixOMath.BuildUp();

            document.SaveAs2(Path.Combine(Environment.CurrentDirectory, "MatrixTypesettingSample.docx"));
            document.Close();

            Console.WriteLine("   ✓ 矩阵排版示例完成，文件保存为: MatrixTypesettingSample.docx");
        }

        /// <summary>
        /// 方程组示例
        /// </summary>
        static async Task EquationSystemSample()
        {
            Console.WriteLine("5. 方程组示例");

            using var application = WordFactory.BlankDocument();
            application.Visible = false;

            IWordDocument document = application.Documents.Add();

            // 添加标题
            IWordRange titleRange = document.Content;
            titleRange.Text = "方程组示例\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;

            // 创建方程组
            var range = document.Range(document.Content.End - 1, document.Content.End);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            IWordOMaths oMaths = range.OMaths;
            IWordRange eqArrayRange = oMaths.Add(range);
            IWordOMath eqArrayOMath = oMaths[1];  // COM集合索引从1开始

            // 创建方程组数组
            var eqArrayFunction = eqArrayOMath.Functions.Add(eqArrayRange, WdOMathFunctionType.wdOMathFunctionEqArray);
            var eqArray = eqArrayFunction.EqArray;

            // 这里简化处理，实际应该通过更复杂的方式添加多行方程
            eqArrayOMath.Range.Text = "{\n" +
                                    "  x + y = 5,\n" +
                                    "  x - y = 1,\n" +
                                    "  2x + 3y = 12\n" +
                                    "}";

            // 设置对齐方式
            eqArray.Align = WdOMathVertAlignType.wdOMathVertAlignCenter;
            eqArray.RowSpacing = 10;

            // 构建专业格式
            eqArrayOMath.BuildUp();

            document.SaveAs2(Path.Combine(Environment.CurrentDirectory, "EquationSystemSample.docx"));
            document.Close();

            Console.WriteLine("   ✓ 方程组示例完成，文件保存为: EquationSystemSample.docx");
        }

        /// <summary>
        /// 嵌套公式示例
        /// </summary>
        static async Task NestedEquationSample()
        {
            Console.WriteLine("6. 嵌套公式示例");

            using var application = WordFactory.BlankDocument();
            application.Visible = false;

            IWordDocument document = application.Documents.Add();

            // 添加标题
            IWordRange titleRange = document.Content;
            titleRange.Text = "嵌套公式示例\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;

            // 创建嵌套公式
            var range = document.Range(document.Content.End - 1, document.Content.End);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            IWordOMaths oMaths = range.OMaths;
            IWordRange nestedRange = oMaths.Add(range);
            IWordOMath nestedOMath = oMaths[1];  // COM集合索引从1开始

            // 创建外层分数
            var outerFractionFunction = nestedOMath.Functions.Add(nestedRange, WdOMathFunctionType.wdOMathFunctionFrac);
            var outerFraction = outerFractionFunction.Frac;

            // 在分子中创建嵌套的平方根
            var innerRadicalFunction = outerFraction.Num.Functions.Add(outerFraction.Num.Range, WdOMathFunctionType.wdOMathFunctionRad);
            var innerRadical = innerRadicalFunction.Rad;
            innerRadical.E.Range.Text = "x^2 + y^2";

            // 设置分母
            outerFraction.Den.Range.Text = "2";

            // 构建专业格式
            nestedOMath.BuildUp();

            document.SaveAs2(Path.Combine(Environment.CurrentDirectory, "NestedEquationSample.docx"));
            document.Close();

            Console.WriteLine("   ✓ 嵌套公式示例完成，文件保存为: NestedEquationSample.docx");
        }

        /// <summary>
        /// 公式样式和格式控制示例
        /// </summary>
        static async Task EquationFormattingSample()
        {
            Console.WriteLine("7. 公式样式和格式控制示例");

            using var application = WordFactory.BlankDocument();
            application.Visible = false;

            IWordDocument document = application.Documents.Add();

            // 添加标题
            IWordRange titleRange = document.Content;
            titleRange.Text = "公式样式和格式控制示例\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;

            // 创建公式
            var range = document.Range(document.Content.End - 1, document.Content.End);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            IWordOMaths oMaths = range.OMaths;
            IWordRange formattedRange = oMaths.Add(range);
            IWordOMath formattedOMath = oMaths[1];  // COM集合索引从1开始

            // 设置公式内容
            formattedOMath.Range.Text = "E = mc^2";

            // 设置公式样式
            formattedOMath.Range.Font.Name = "Times New Roman";
            formattedOMath.Range.Font.Size = 14;
            formattedOMath.Range.Font.Bold = false;
            formattedOMath.Range.Font.Color = WdColor.wdColorBlue;

            // 居中对齐
            formattedOMath.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            // 设置公式类型为专业显示格式
            formattedOMath.Type = WdOMathType.wdOMathDisplay;
            formattedOMath.Justification = WdOMathJc.wdOMathJcCenter;

            // 构建专业格式
            formattedOMath.BuildUp();

            // 添加样式说明
            var descRange = document.Range(document.Content.End - 1, document.Content.End);
            descRange.InsertParagraphAfter();
            descRange.Text = "• 字体: Times New Roman, 14pt\n• 颜色: 蓝色\n• 对齐: 居中\n• 格式: 专业显示";

            document.SaveAs2(Path.Combine(Environment.CurrentDirectory, "EquationFormattingSample.docx"));
            document.Close();

            Console.WriteLine("   ✓ 公式样式和格式控制示例完成，文件保存为: EquationFormattingSample.docx");
        }

        /// <summary>
        /// LaTeX转Word公式示例
        /// </summary>
        static async Task LaTeXToWordSample()
        {
            Console.WriteLine("8. LaTeX转Word公式示例");

            using var application = WordFactory.BlankDocument();
            application.Visible = false;

            IWordDocument document = application.Documents.Add();

            // 添加标题
            IWordRange titleRange = document.Content;
            titleRange.Text = "LaTeX转Word公式示例\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;

            // LaTeX公式示例
            var latexFormulas = new[]
            {
                "\\frac{d^2y}{dx^2} + \\omega^2 y = 0",
                "\\int_{0}^{\\infty} e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}",
                "\\sum_{i=1}^{n} i^2 = \\frac{n(n+1)(2n+1)}{6}"
            };

            var converter = new LaTeXToWordConverter();

            foreach (string latexFormula in latexFormulas)
            {
                // 插入LaTeX原文
                var latexRange = document.Range(document.Content.End - 1, document.Content.End);
                latexRange.InsertParagraphAfter();
                latexRange.Text = $"LaTeX: {latexFormula}";
                latexRange.Font.Italic = true;

                // 转换并插入Word公式
                var range = document.Range(document.Content.End - 1, document.Content.End);
                range.InsertParagraphAfter();
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                IWordOMath oMath = converter.ConvertLaTeXToWordFormula(range, latexFormula);
                oMath.BuildUp();
                oMath.Type = WdOMathType.wdOMathDisplay;
            }

            document.SaveAs2(Path.Combine(Environment.CurrentDirectory, "LaTeXToWordSample.docx"));
            document.Close();

            Console.WriteLine("   ✓ LaTeX转Word公式示例完成，文件保存为: LaTeXToWordSample.docx");
        }

        /// <summary>
        /// 学术论文自动化排版示例
        /// </summary>
        static async Task ScientificPaperSample()
        {
            Console.WriteLine("9. 学术论文自动化排版示例");

            // 模拟学术论文的LaTeX公式
            var paperEquations = new List<string>
            {
                "\\frac{\\partial^2 u}{\\partial t^2} = c^2 \\nabla^2 u",  // 波动方程
                "\\int_0^L \\rho(x) dx = M",                                // 质量积分
                "\\begin{pmatrix} \\cos\\theta & -\\sin\\theta \\\\ \\sin\\theta & \\cos\\theta \\end{pmatrix}",  // 旋转矩阵
                "\\lim_{n \\to \\infty} \\left(1 + \\frac{1}{n}\\right)^n = e"  // 重要极限
            };

            var formatter = new ScientificPaperFormatter();
            formatter.FormatScientificPaper(
                Path.Combine(Environment.CurrentDirectory, "PaperTemplate.docx"),
                paperEquations,
                Path.Combine(Environment.CurrentDirectory, "ScientificPaperSample.docx")
            );

            Console.WriteLine("   ✓ 学术论文自动化排版示例完成，文件保存为: ScientificPaperSample.docx");
        }
    }
}