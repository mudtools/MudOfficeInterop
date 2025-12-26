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
    /// 公式处理器
    /// </summary>
    public class EquationProcessor
    {
        private readonly LaTeXToWordConverter _converter;
        private int _equationCounter = 0;

        public EquationProcessor()
        {
            _converter = new LaTeXToWordConverter();
        }

        /// <summary>
        /// 处理所有公式
        /// </summary>
        /// <param name="document">Word文档</param>
        /// <param name="latexEquations">LaTeX公式列表</param>
        public void ProcessEquations(IWordDocument document, List<string> latexEquations)
        {
            // 创建公式样式
            IWordStyle equationStyle = CreateEquationStyle(document);

            foreach (string latexEquation in latexEquations)
            {
                // 插入新段落
                IWordRange insertRange = document.Range(document.Content.End - 1, document.Content.End);
                insertRange.InsertParagraphAfter();
                insertRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 转换并插入公式
                IWordOMath oMath = _converter.ConvertLaTeXToWordFormula(insertRange, latexEquation);

                // 应用样式和编号
                ApplyEquationFormatting(oMath, equationStyle);
                AddEquationNumber(oMath, ++_equationCounter);
            }
        }

        /// <summary>
        /// 创建公式样式
        /// </summary>
        private IWordStyle CreateEquationStyle(IWordDocument document)
        {
            IWordStyle equationStyle;
            try
            {
                // 尝试获取现有样式
                equationStyle = document.Styles["Equation"];
            }
            catch
            {
                // 样式不存在，创建新样式
                equationStyle = document.Styles.Add("Equation", WdStyleType.wdStyleTypeParagraph);
            }

            equationStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            equationStyle.ParagraphFormat.SpaceAfter = 12;
            equationStyle.ParagraphFormat.SpaceBefore = 12;
            equationStyle.Font.Name = "Times New Roman";
            equationStyle.Font.Size = 12;

            return equationStyle;
        }

        /// <summary>
        /// 应用公式格式
        /// </summary>
        private void ApplyEquationFormatting(IWordOMath oMath, IWordStyle style)
        {
            // 应用段落样式
            oMath.Range.Style = style;

            // 设置公式类型为专业显示格式
            oMath.Type = WdOMathType.wdOMathDisplay;
            oMath.Justification = WdOMathJc.wdOMathJcCenter;

            // 构建专业格式
            oMath.BuildUp();
        }

        /// <summary>
        /// 添加公式编号
        /// </summary>
        private void AddEquationNumber(IWordOMath oMath, int number)
        {
            // 在公式后添加编号
            IWordRange endRange = oMath.Range.Duplicate;
            endRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            endRange.Text = $"    ({number})";

            // 创建书签
            string bookmarkName = $"EQ_{number}";
            try
            {
                oMath.Range.Document.Bookmarks.Add(bookmarkName, oMath.Range);
            }
            catch
            {
                // 书签可能已存在，忽略错误
            }
        }
    }

    /// <summary>
    /// 期刊样式管理器
    /// </summary>
    public class JournalStyleManager
    {
        /// <summary>
        /// 应用期刊模板
        /// </summary>
        /// <param name="document">Word文档</param>
        /// <param name="journalName">期刊名称</param>
        public void ApplyJournalTemplate(IWordDocument document, string journalName)
        {
            switch (journalName.ToLower())
            {
                case "ieee":
                    ApplyIEEEStyle(document);
                    break;
                case "nature":
                    ApplyNatureStyle(document);
                    break;
                case "science":
                    ApplyScienceStyle(document);
                    break;
                default:
                    ApplyDefaultStyle(document);
                    break;
            }
        }

        /// <summary>
        /// 应用IEEE期刊样式
        /// </summary>
        private void ApplyIEEEStyle(IWordDocument document)
        {
            // IEEE期刊的公式样式要求
            foreach (IWordOMath oMath in document.OMaths)
            {
                oMath.Range.Font.Name = "Times New Roman";
                oMath.Range.Font.Size = 10;  // IEEE要求较小字号
                oMath.Range.ParagraphFormat.SpaceAfter = 6;
                oMath.Range.ParagraphFormat.SpaceBefore = 6;
            }

            // 设置文档整体格式
            document.Content.Font.Name = "Times New Roman";
            document.Content.Font.Size = 10;
        }

        /// <summary>
        /// 应用Nature期刊样式
        /// </summary>
        private void ApplyNatureStyle(IWordDocument document)
        {
            foreach (IWordOMath oMath in document.OMaths)
            {
                oMath.Range.Font.Name = "Times New Roman";
                oMath.Range.Font.Size = 12;
                oMath.Range.ParagraphFormat.SpaceAfter = 10;
                oMath.Range.ParagraphFormat.SpaceBefore = 10;
            }
        }

        /// <summary>
        /// 应用Science期刊样式
        /// </summary>
        private void ApplyScienceStyle(IWordDocument document)
        {
            foreach (IWordOMath oMath in document.OMaths)
            {
                oMath.Range.Font.Name = "Times New Roman";
                oMath.Range.Font.Size = 11;
                oMath.Range.ParagraphFormat.SpaceAfter = 8;
                oMath.Range.ParagraphFormat.SpaceBefore = 8;
            }
        }

        /// <summary>
        /// 应用默认样式
        /// </summary>
        private void ApplyDefaultStyle(IWordDocument document)
        {
            foreach (IWordOMath oMath in document.OMaths)
            {
                oMath.Range.Font.Name = "Times New Roman";
                oMath.Range.Font.Size = 12;
                oMath.Range.ParagraphFormat.SpaceAfter = 12;
                oMath.Range.ParagraphFormat.SpaceBefore = 12;
            }
        }
    }

    /// <summary>
    /// 科学论文格式化器
    /// </summary>
    public class ScientificPaperFormatter
    {
        /// <summary>
        /// 格式化科学论文
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="equations">公式列表</param>
        /// <param name="outputPath">输出路径</param>
        public void FormatScientificPaper(string templatePath, List<string> equations, string outputPath)
        {
            using var application = WordFactory.BlankDocument();
            application.Visible = false;

            IWordDocument document;

            // 尝试打开模板，如果不存在则创建新文档
            try
            {
                if (File.Exists(templatePath))
                {
                    document = application.Documents.Open(templatePath);
                }
                else
                {
                    document = application.Documents.Add();
                    CreatePaperTemplate(document);
                }
            }
            catch
            {
                // 如果模板文件损坏，创建新文档
                document = application.Documents.Add();
                CreatePaperTemplate(document);
            }

            try
            {
                // 初始化处理器
                var processor = new EquationProcessor();
                var styleManager = new JournalStyleManager();

                // 添加论文标题
                AddPaperTitle(document, "科学论文公式排版示例");

                // 处理所有公式
                processor.ProcessEquations(document, equations);

                // 添加公式目录
                AddEquationList(document, equations);

                // 应用期刊样式
                styleManager.ApplyJournalTemplate(document, "IEEE");

                // 保存文档
                document.SaveAs2(outputPath);
            }
            finally
            {
                document.Close();
                application.Quit();
            }
        }

        /// <summary>
        /// 创建论文模板
        /// </summary>
        private void CreatePaperTemplate(IWordDocument document)
        {
            // 添加基本文档结构
            var titleRange = document.Content;
            titleRange.Text = "科学论文模板\n\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;

            var abstractRange = document.Range(document.Content.End - 1, document.Content.End);
            abstractRange.Text = "摘要：\n这里是论文摘要...\n\n";
            abstractRange.Font.Bold = true;
            abstractRange.Font.Size = 12;

            var contentRange = document.Range(document.Content.End - 1, document.Content.End);
            contentRange.Text = "1. 引言\n这里是引言内容...\n\n";
            contentRange.Text += "2. 主要内容\n\n";
            contentRange.Font.Size = 12;
        }

        /// <summary>
        /// 添加论文标题
        /// </summary>
        private void AddPaperTitle(IWordDocument document, string title)
        {
            var titleRange = document.Range(0, 0);
            titleRange.Text = $"{title}\n\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 18;
            titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        /// <summary>
        /// 添加公式列表
        /// </summary>
        private void AddEquationList(IWordDocument document, List<string> equations)
        {
            var listRange = document.Range(document.Content.End - 1, document.Content.End);
            listRange.InsertParagraphAfter();
            listRange.Text = "\n公式列表：\n";
            listRange.Font.Bold = true;

            for (int i = 0; i < equations.Count; i++)
            {
                var eqRange = document.Range(document.Content.End - 1, document.Content.End);
                eqRange.InsertParagraphAfter();
                eqRange.Text = $"公式 ({i + 1}): {equations[i]}";
                eqRange.Font.Italic = true;
            }
        }
    }
}