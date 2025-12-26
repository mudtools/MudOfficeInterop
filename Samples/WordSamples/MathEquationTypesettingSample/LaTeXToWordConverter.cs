//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace MathEquationTypesettingSample
{
    /// <summary>
    /// LaTeX公式元素类型
    /// </summary>
    public enum LaTeXType
    {
        Text,
        Fraction,
        Matrix,
        Integral,
        Summation,
        Radical,
        Superscript,
        Subscript,
        SubSuperscript,
        Function,
        Symbol
    }

    /// <summary>
    /// LaTeX公式元素
    /// </summary>
    public class LaTeXElement
    {
        public LaTeXType Type { get; set; }
        public string Content { get; set; } = string.Empty;
        public List<LaTeXElement> Children { get; set; } = new List<LaTeXElement>();
        public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>();
    }

    /// <summary>
    /// LaTeX到Word公式转换器
    /// </summary>
    public class LaTeXToWordConverter
    {
        /// <summary>
        /// 将LaTeX公式转换为Word数学对象
        /// </summary>
        /// <param name="range">插入位置</param>
        /// <param name="latexFormula">LaTeX公式字符串</param>
        /// <returns>Word数学对象</returns>
        public IWordOMath ConvertLaTeXToWordFormula(IWordRange range, string latexFormula)
        {
            // 解析LaTeX公式
            var parsedFormula = ParseLaTeX(latexFormula);

            // 创建Word公式
            IWordOMaths oMaths = range.OMaths;
            IWordRange formulaRange = oMaths.Add(range);
            IWordOMath oMath = oMaths[1];  // COM集合索引从1开始

            // 递归构建公式结构
            BuildFormulaStructure(oMath, parsedFormula);

            return oMath;
        }

        /// <summary>
        /// 解析LaTeX公式字符串
        /// </summary>
        /// <param name="latexFormula">LaTeX公式</param>
        /// <returns>解析后的元素树</returns>
        public LaTeXElement ParseLaTeX(string latexFormula)
        {
            var root = new LaTeXElement { Type = LaTeXType.Text, Content = latexFormula };

            // 简化的LaTeX解析逻辑
            if (latexFormula.Contains(@"\frac"))
            {
                root = ParseFraction(latexFormula);
            }
            else if (latexFormula.Contains(@"\int") || latexFormula.Contains(@"\sum"))
            {
                root = ParseNary(latexFormula);
            }
            else if (latexFormula.Contains(@"\sqrt") || latexFormula.Contains(@"\begin{pmatrix}"))
            {
                if (latexFormula.Contains(@"\begin{pmatrix}"))
                {
                    root = ParseMatrix(latexFormula);
                }
                else
                {
                    root = ParseRadical(latexFormula);
                }
            }
            else if (latexFormula.Contains("^"))
            {
                root = ParseSuperscript(latexFormula);
            }

            return root;
        }

        /// <summary>
        /// 解析分数
        /// </summary>
        private LaTeXElement ParseFraction(string fraction)
        {
            var element = new LaTeXElement { Type = LaTeXType.Fraction };

            // 简化的分数解析，提取分子和分母
            var match = System.Text.RegularExpressions.Regex.Match(fraction, @"\\frac\{([^}]*)\}\{([^}]*)\}");
            if (match.Success)
            {
                element.Children.Add(new LaTeXElement
                {
                    Type = LaTeXType.Text,
                    Content = match.Groups[1].Value
                }); // 分子
                element.Children.Add(new LaTeXElement
                {
                    Type = LaTeXType.Text,
                    Content = match.Groups[2].Value
                }); // 分母
            }

            return element;
        }

        /// <summary>
        /// 解析n元运算符（积分、求和等）
        /// </summary>
        private LaTeXElement ParseNary(string nary)
        {
            var element = new LaTeXElement { Type = LaTeXType.Integral };

            if (nary.Contains(@"\int"))
            {
                element.Type = LaTeXType.Integral;
                element.Content = "∫";
            }
            else if (nary.Contains(@"\sum"))
            {
                element.Type = LaTeXType.Summation;
                element.Content = "∑";
            }

            // 解析表达式、上标和下标
            var expressionMatch = System.Text.RegularExpressions.Regex.Match(nary, @"[=]?\s*([^{]+)");
            if (expressionMatch.Success)
            {
                element.Children.Add(new LaTeXElement
                {
                    Type = LaTeXType.Text,
                    Content = expressionMatch.Groups[1].Value.Trim()
                });
            }

            // 解析上下标（简化处理）
            var subMatch = System.Text.RegularExpressions.Regex.Match(nary, @"_{([^}]*)");
            if (subMatch.Success)
            {
                element.Attributes["sub"] = subMatch.Groups[1].Value;
            }

            var supMatch = System.Text.RegularExpressions.Regex.Match(nary, @"\^{([^}]*)");
            if (supMatch.Success)
            {
                element.Attributes["sup"] = supMatch.Groups[1].Value;
            }

            return element;
        }

        /// <summary>
        /// 解析矩阵
        /// </summary>
        private LaTeXElement ParseMatrix(string matrix)
        {
            var element = new LaTeXElement { Type = LaTeXType.Matrix };

            // 提取矩阵内容
            var contentMatch = System.Text.RegularExpressions.Regex.Match(matrix, @"\\begin\{pmatrix\}(.*?)\\end\{pmatrix\}",
                System.Text.RegularExpressions.RegexOptions.Singleline);

            if (contentMatch.Success)
            {
                var content = contentMatch.Groups[1].Value;
                var rows = content.Split(new[] { @"\\" }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var row in rows)
                {
                    var rowElement = new LaTeXElement { Type = LaTeXType.Text };
                    var elements = row.Split('&');
                    foreach (var cell in elements)
                    {
                        rowElement.Children.Add(new LaTeXElement
                        {
                            Type = LaTeXType.Text,
                            Content = cell.Trim()
                        });
                    }
                    element.Children.Add(rowElement);
                }
            }

            return element;
        }

        /// <summary>
        /// 解析根式
        /// </summary>
        private LaTeXElement ParseRadical(string radical)
        {
            var element = new LaTeXElement { Type = LaTeXType.Radical };

            var match = System.Text.RegularExpressions.Regex.Match(radical, @"\\sqrt\[?([^]\}]*)\]?\{([^}]*)\}");
            if (match.Success)
            {
                if (!string.IsNullOrEmpty(match.Groups[1].Value))
                {
                    element.Attributes["degree"] = match.Groups[1].Value; // 根次
                }
                element.Children.Add(new LaTeXElement
                {
                    Type = LaTeXType.Text,
                    Content = match.Groups[2].Value
                }); // 被开方数
            }
            else
            {
                // 简单的平方根
                var simpleMatch = System.Text.RegularExpressions.Regex.Match(radical, @"\\sqrt\{([^}]*)\}");
                if (simpleMatch.Success)
                {
                    element.Children.Add(new LaTeXElement
                    {
                        Type = LaTeXType.Text,
                        Content = simpleMatch.Groups[1].Value
                    });
                }
            }

            return element;
        }

        /// <summary>
        /// 解析上标
        /// </summary>
        private LaTeXElement ParseSuperscript(string superscript)
        {
            var element = new LaTeXElement { Type = LaTeXType.Superscript };

            var match = System.Text.RegularExpressions.Regex.Match(superscript, @"([a-zA-Z]+)\^{([^}]*)");
            if (match.Success)
            {
                element.Children.Add(new LaTeXElement
                {
                    Type = LaTeXType.Text,
                    Content = match.Groups[1].Value
                }); // 基底
                element.Children.Add(new LaTeXElement
                {
                    Type = LaTeXType.Text,
                    Content = match.Groups[2].Value
                }); // 上标
            }

            return element;
        }

        /// <summary>
        /// 构建Word公式结构
        /// </summary>
        private void BuildFormulaStructure(IWordOMath oMath, LaTeXElement element)
        {
            switch (element.Type)
            {
                case LaTeXType.Fraction:
                    CreateFraction(oMath, element);
                    break;
                case LaTeXType.Matrix:
                    CreateMatrix(oMath, element);
                    break;
                case LaTeXType.Integral:
                case LaTeXType.Summation:
                    CreateNary(oMath, element);
                    break;
                case LaTeXType.Radical:
                    CreateRadical(oMath, element);
                    break;
                case LaTeXType.Superscript:
                    CreateSuperscript(oMath, element);
                    break;
                default:
                    // 普通文本
                    oMath.Range.Text = element.Content;
                    break;
            }
        }

        /// <summary>
        /// 创建分数
        /// </summary>
        private void CreateFraction(IWordOMath oMath, LaTeXElement element)
        {
            var fractionFunction = oMath.Functions.Add(oMath.Range, WdOMathFunctionType.wdOMathFunctionFrac);
            var fraction = fractionFunction.Frac;

            if (element.Children.Count >= 2)
            {
                fraction.Num.Range.Text = element.Children[0].Content;
                fraction.Den.Range.Text = element.Children[1].Content;
            }
        }

        /// <summary>
        /// 创建矩阵
        /// </summary>
        private void CreateMatrix(IWordOMath oMath, LaTeXElement element)
        {
            var matrixFunction = oMath.Functions.Add(oMath.Range, WdOMathFunctionType.wdOMathFunctionMat);
            var matrix = matrixFunction.Mat;

            // 添加行
            for (int i = 0; i < element.Children.Count; i++)
            {
                matrix.Rows.Add(null);
            }

            // 添加列
            if (element.Children.Count > 0)
            {
                int maxCols = element.Children.Max(row => row.Children.Count);
                for (int j = 0; j < maxCols; j++)
                {
                    matrix.Cols.Add(null);
                }

                // 填充矩阵元素
                for (int row = 0; row < element.Children.Count; row++)
                {
                    for (int col = 0; col < element.Children[row].Children.Count; col++)
                    {
                        matrix.Cell(row + 1, col + 1).Range.Text = element.Children[row].Children[col].Content;
                    }
                }
            }
        }

        /// <summary>
        /// 创建n元运算符（积分、求和等）
        /// </summary>
        private void CreateNary(IWordOMath oMath, LaTeXElement element)
        {
            var naryFunction = oMath.Functions.Add(oMath.Range, WdOMathFunctionType.wdOMathFunctionNary);
            var nary = naryFunction.Nary;

            nary.Char = element.Content[0];

            if (element.Attributes.ContainsKey("sub"))
            {
                nary.Sub.Range.Text = element.Attributes["sub"];
            }

            if (element.Attributes.ContainsKey("sup"))
            {
                nary.Sup.Range.Text = element.Attributes["sup"];
            }

            if (element.Children.Count > 0)
            {
                nary.E.Range.Text = element.Children[0].Content;
            }
        }

        /// <summary>
        /// 创建根式
        /// </summary>
        private void CreateRadical(IWordOMath oMath, LaTeXElement element)
        {
            var radicalFunction = oMath.Functions.Add(oMath.Range, WdOMathFunctionType.wdOMathFunctionRad);
            var radical = radicalFunction.Rad;

            if (element.Attributes.ContainsKey("degree"))
            {
                radical.Deg.Range.Text = element.Attributes["degree"];
            }

            if (element.Children.Count > 0)
            {
                radical.E.Range.Text = element.Children[0].Content;
            }
        }

        /// <summary>
        /// 创建上标
        /// </summary>
        private void CreateSuperscript(IWordOMath oMath, LaTeXElement element)
        {
            var supFunction = oMath.Functions.Add(oMath.Range, WdOMathFunctionType.wdOMathFunctionScrSup);
            var sup = supFunction.ScrSup;

            if (element.Children.Count >= 2)
            {
                sup.E.Range.Text = element.Children[0].Content;
                sup.Sup.Range.Text = element.Children[1].Content;
            }
        }
    }
}