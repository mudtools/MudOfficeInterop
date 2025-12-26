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
    /// Word操作辅助类
    /// 提供简化和安全的Word对象操作方法
    /// </summary>
    public static class WordHelper
    {

        public static IWordRange GetEndRange(IWordDocument document)
        {
            var range1 = document.Content;
            //range1.Collapse(WdCollapseDirection.wdCollapseEnd);
            return range1;
        }

        /// <summary>
        /// 安全地添加数学公式并返回新创建的公式对象
        /// </summary>
        /// <param name="range">插入位置的范围</param>
        /// <returns>新创建的数学公式对象</returns>
        public static IWordOMath AddMathEquation(IWordRange range)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));

            IWordOMaths oMaths = range.OMaths;
            var returnRange = oMaths.Add(range);

            // 返回刚添加的公式
            // COM集合索引从1开始，新添加的公式是集合中的最后一个
            return oMaths[oMaths.Count];
        }

        /// <summary>
        /// 添加数学公式并设置文本内容
        /// </summary>
        /// <param name="range">插入位置的范围</param>
        /// <param name="text">公式文本内容</param>
        /// <returns>新创建的数学公式对象</returns>
        public static IWordOMath AddMathEquation(IWordRange range, string text)
        {
            IWordOMath oMath = AddMathEquation(range);
            oMath.Range.Text = text;
            return oMath;
        }

        /// <summary>
        /// 创建分数公式
        /// </summary>
        /// <param name="range">插入位置的范围</param>
        /// <param name="numerator">分子文本</param>
        /// <param name="denominator">分母文本</param>
        /// <param name="fractionType">分数类型</param>
        /// <returns>创建的分数公式</returns>
        public static IWordOMath CreateFraction(IWordRange range, string numerator, string denominator,
            WdOMathFracType fractionType = WdOMathFracType.wdOMathFracBar)
        {
            IWordOMath oMath = AddMathEquation(range);
            var fractionFunction = oMath.Functions.Add(oMath.Range, WdOMathFunctionType.wdOMathFunctionFrac);
            var fraction = fractionFunction.Frac;

            if (!string.IsNullOrEmpty(numerator))
                fraction.Num.Range.Text = numerator;

            if (!string.IsNullOrEmpty(denominator))
                fraction.Den.Range.Text = denominator;

            fraction.Type = fractionType;
            return oMath;
        }

        /// <summary>
        /// 创建根式公式
        /// </summary>
        /// <param name="range">插入位置的范围</param>
        /// <param name="radicand">被开方数</param>
        /// <param name="degree">根次（可选）</param>
        /// <returns>创建的根式公式</returns>
        public static IWordOMath CreateRadical(IWordRange range, string radicand, string? degree = null)
        {
            IWordOMath oMath = AddMathEquation(range);
            var radicalFunction = oMath.Functions.Add(oMath.Range, WdOMathFunctionType.wdOMathFunctionRad);
            var radical = radicalFunction.Rad;

            if (!string.IsNullOrEmpty(radicand))
                radical.E.Range.Text = radicand;

            if (!string.IsNullOrEmpty(degree))
                radical.Deg.Range.Text = degree;

            return oMath;
        }

        /// <summary>
        /// 创建n元运算符（积分、求和等）
        /// </summary>
        /// <param name="range">插入位置的范围</param>
        /// <param name="character">运算符字符</param>
        /// <param name="expression">表达式</param>
        /// <param name="subscript">下标（可选）</param>
        /// <param name="superscript">上标（可选）</param>
        /// <returns>创建的n元运算符公式</returns>
        public static IWordOMath CreateNary(IWordRange range, string character, string expression,
            string? subscript = null, string? superscript = null)
        {
            IWordOMath oMath = AddMathEquation(range);
            var naryFunction = oMath.Functions.Add(oMath.Range, WdOMathFunctionType.wdOMathFunctionNary);
            var nary = naryFunction.Nary;

            nary.Char = character[0];

            if (!string.IsNullOrEmpty(expression))
                nary.E.Range.Text = expression;

            if (!string.IsNullOrEmpty(subscript))
                nary.Sub.Range.Text = subscript;

            if (!string.IsNullOrEmpty(superscript))
                nary.Sup.Range.Text = superscript;

            return oMath;
        }

        /// <summary>
        /// 创建上标公式
        /// </summary>
        /// <param name="range">插入位置的范围</param>
        /// <param name="baseText">基底文本</param>
        /// <param name="superscriptText">上标文本</param>
        /// <returns>创建的上标公式</returns>
        public static IWordOMath CreateSuperscript(IWordRange range, string baseText, string superscriptText)
        {
            IWordOMath oMath = AddMathEquation(range);
            var supFunction = oMath.Functions.Add(oMath.Range, WdOMathFunctionType.wdOMathFunctionScrSup);
            var sup = supFunction.ScrSup;

            if (!string.IsNullOrEmpty(baseText))
                sup.E.Range.Text = baseText;

            if (!string.IsNullOrEmpty(superscriptText))
                sup.Sup.Range.Text = superscriptText;

            return oMath;
        }

        /// <summary>
        /// 安全地构建公式专业格式
        /// </summary>
        /// <param name="oMath">数学公式对象</param>
        public static void BuildUpSafely(IWordOMath oMath)
        {
            if (oMath != null && oMath.Range != null)
            {
                try
                {
                    oMath.BuildUp();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"构建公式格式时出错: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 设置公式基本格式
        /// </summary>
        /// <param name="oMath">数学公式对象</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="alignment">对齐方式</param>
        /// <param name="equationType">公式类型</param>
        public static void SetEquationFormat(IWordOMath oMath,
            string? fontName = null,
            int? fontSize = null,
            WdParagraphAlignment? alignment = null,
            WdOMathType? equationType = null)
        {
            if (oMath?.Range == null) return;

            if (!string.IsNullOrEmpty(fontName))
                oMath.Range.Font.Name = fontName;

            if (fontSize.HasValue)
                oMath.Range.Font.Size = fontSize.Value;

            if (alignment.HasValue)
                oMath.Range.ParagraphFormat.Alignment = alignment.Value;

            if (equationType.HasValue)
                oMath.Type = equationType.Value;
        }

        /// <summary>
        /// 在指定位置插入段落并返回新的范围
        /// </summary>
        /// <param name="document">文档对象</param>
        /// <param name="atEnd">是否在文档末尾插入</param>
        /// <returns>新段落的范围</returns>
        public static IWordRange InsertNewParagraph(IWordDocument document, bool atEnd = true)
        {
            IWordRange range;
            if (atEnd)
            {
                range = document.Range(document.Content.End - 1, document.Content.End);
            }
            else
            {
                range = document.Range(0, 0);
            }

            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            return range;
        }

        /// <summary>
        /// 批量设置公式格式
        /// </summary>
        /// <param name="document">文档对象</param>
        /// <param name="action">格式设置操作</param>
        public static void ApplyToAllEquations(IWordDocument document, Action<IWordOMath> action)
        {
            if (document?.OMaths == null) return;

            foreach (IWordOMath oMath in document.OMaths)
            {
                try
                {
                    action(oMath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"处理公式时出错: {ex.Message}");
                }
            }
        }
    }
}