//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace TextFormattingSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 文本格式化示例");

            // 示例1: 字体格式设置
            Console.WriteLine("\n=== 示例1: 字体格式设置 ===");
            FontFormattingDemo();

            // 示例2: 段落格式设置
            Console.WriteLine("\n=== 示例2: 段落格式设置 ===");
            ParagraphFormattingDemo();

            // 示例3: 样式应用
            Console.WriteLine("\n=== 示例3: 样式应用 ===");
            StyleApplicationDemo();

            // 示例4: 列表和编号
            Console.WriteLine("\n=== 示例4: 列表和编号 ===");
            ListAndNumberingDemo();

            // 示例5: 边框和底纹
            Console.WriteLine("\n=== 示例5: 边框和底纹 ===");
            BordersAndShadingDemo();

            // 示例6: 制表符设置
            Console.WriteLine("\n=== 示例6: 制表符设置 ===");
            TabSettingsDemo();

            // 示例7: 字符和段落间距
            Console.WriteLine("\n=== 示例7: 字符和段落间距 ===");
            SpacingSettingsDemo();

            // 示例8: 文本高亮和下划线
            Console.WriteLine("\n=== 示例8: 文本高亮和下划线 ===");
            TextHighlightingAndUnderlineDemo();

            // 示例9: 字符缩放和间距
            Console.WriteLine("\n=== 示例9: 字符缩放和间距 ===");
            CharacterScalingAndSpacingDemo();

            // 示例10: 文本效果
            Console.WriteLine("\n=== 示例10: 文本效果 ===");
            TextEffectsDemo();

            // 示例11: 多级标题和目录
            Console.WriteLine("\n=== 示例11: 多级标题和目录 ===");
            MultiLevelHeadingsAndTOCDemo();

            // 示例12: 实际应用示例
            Console.WriteLine("\n=== 示例12: 实际应用示例 ===");
            RealWorldApplicationDemo();

            // 示例13: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例13: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 字体格式设置示例
        /// </summary>
        static void FontFormattingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 获取文档范围
                var range = document.Range();

                // 添加标题文本
                range.Text = "文档标题\n";

                range.Font.Name = "微软雅黑";
                range.Font.Size = 18;
                range.Font.Bold = true;
                range.Font.Color = WdColor.wdColorBlue;

                // 添加正文内容
                var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                contentRange.Text = "这是文档的正文内容，使用标准字体格式。\n";
                contentRange.Font.Name = "宋体";
                contentRange.Font.Size = 12;
                contentRange.Font.Bold = false;
                contentRange.Font.Italic = false;

                Console.WriteLine("字体格式设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"字体格式设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 段落格式设置示例
        /// </summary>
        static void ParagraphFormattingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加标题
                var titleRange = document.Range();
                titleRange.Text = "居中标题\n";
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 12;

                // 添加左对齐段落
                var leftPara = document.Range(document.Content.End - 1, document.Content.End - 1);
                leftPara.Text = "这是左对齐的段落文本。\n";
                leftPara.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                // 添加右对齐段落
                var rightPara = document.Range(document.Content.End - 1, document.Content.End - 1);
                rightPara.Text = "这是右对齐的段落文本。\n";
                rightPara.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                // 添加两端对齐段落
                var justifyPara = document.Range(document.Content.End - 1, document.Content.End - 1);
                justifyPara.Text = "这是两端对齐的段落文本，文本会自动调整以填满整行。\n";
                justifyPara.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                Console.WriteLine("段落格式设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"段落格式设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 样式应用示例
        /// </summary>
        static void StyleApplicationDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 应用内置样式
                var heading1 = document.Range();
                heading1.Text = "标题 1\n";
                heading1.Style = "标题 1";

                var heading2 = document.Range(document.Content.End - 1, document.Content.End - 1);
                heading2.Text = "标题 2\n";
                heading2.Style = "标题 2";

                var normalText = document.Range(document.Content.End - 1, document.Content.End - 1);
                normalText.Text = "正文文本\n";
                normalText.Style = "正文";

                // 创建自定义样式
                var customStyle = document.Styles.Add("我的自定义样式");
                customStyle.Font.Name = "楷体";
                customStyle.Font.Size = 14;
                customStyle.Font.Color = WdColor.wdColorRed;
                customStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                var customRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                customRange.Text = "使用自定义样式的文本\n";
                customRange.Style = "我的自定义样式";

                Console.WriteLine("样式应用完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"样式应用出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 列表和编号示例
        /// </summary>
        static void ListAndNumberingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 创建项目符号列表
                var bulletList = document.Range();
                bulletList.Text = "项目 1\n项目 2\n项目 3\n";
                // 注意：由于库的限制，这里可能需要使用不同的方法来应用列表格式

                // 创建编号列表
                var numberedList = document.Range(document.Content.End - 1, document.Content.End - 1);
                numberedList.Text = "第一项\n第二项\n第三项\n";
                // 注意：由于库的限制，这里可能需要使用不同的方法来应用列表格式

                // 创建多级列表
                var multiLevelList = document.Range(document.Content.End - 1, document.Content.End - 1);
                multiLevelList.Text = "主要项目\n子项目 1\n子项目 2\n另一个主要项目\n其子项目\n";
                // 注意：由于库的限制，这里可能需要使用不同的方法来应用列表格式

                Console.WriteLine("列表和编号设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"列表和编号设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 边框和底纹示例
        /// </summary>
        static void BordersAndShadingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加文本
                var range = document.Range();
                range.Text = "带边框和底纹的文本\n";

                // 设置边框
                range.Borders.Enable = true;
                range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;

                // 设置底纹
                range.Shading.BackgroundPatternColor = WdColor.wdColorLightYellow;

                Console.WriteLine("边框和底纹设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"边框和底纹设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 制表符设置示例
        /// </summary>
        static void TabSettingsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 设置制表符
                var range = document.Range();
                range.Text = "姓名\t年龄\t职业\n张三\t25\t工程师\n李四\t30\t设计师\n";

                // 在特定位置添加制表符
                range.ParagraphFormat.TabStops.Add(100, WdTabAlignment.wdAlignTabLeft);
                range.ParagraphFormat.TabStops.Add(200, WdTabAlignment.wdAlignTabLeft);
                range.ParagraphFormat.TabStops.Add(300, WdTabAlignment.wdAlignTabLeft);

                Console.WriteLine("制表符设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"制表符设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 字符和段落间距设置示例
        /// </summary>
        static void SpacingSettingsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加文本
                var range = document.Range();
                range.Text = "字符和段落间距设置示例文本。\n这是第二段文本用于演示段落间距。\n这是第三段文本。";

                // 设置段落间距
                range.ParagraphFormat.SpaceBefore = 12;  // 段前间距
                range.ParagraphFormat.SpaceAfter = 12;   // 段后间距
                range.ParagraphFormat.LineSpacing = 1.5f; // 行距1.5倍
                Console.WriteLine("段落间距设置完成");

                Console.WriteLine("字符和段落间距设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"字符和段落间距设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文本高亮和下划线示例
        /// </summary>
        static void TextHighlightingAndUnderlineDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加文本
                var range = document.Range();
                range.Text = "文本高亮和下划线示例。\n";

                // 设置文本高亮
                var highlightRange = document.Range(0, 5);
                highlightRange.HighlightColorIndex = WdColorIndex.wdYellow;

                // 设置不同类型的下划线
                var underlineRange = document.Range(6, 10);
                underlineRange.Font.Underline = WdUnderline.wdUnderlineSingle;

                var doubleUnderlineRange = document.Range(11, 15);
                doubleUnderlineRange.Font.Underline = WdUnderline.wdUnderlineDouble;

                Console.WriteLine("文本高亮和下划线设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文本高亮和下划线设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 字符缩放和间距示例
        /// </summary>
        static void CharacterScalingAndSpacingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加文本
                var range = document.Range();
                range.Text = "字符缩放和间距示例文本。\n";

                // 设置字符间距
                var spacingRange = document.Range(5, 9);
                spacingRange.Font.Spacing = 2.0f; // 字符间距2磅

                // 设置字符位置
                var positionRange = document.Range(10, 14);
                positionRange.Font.Position = -6; // 下标位置

                Console.WriteLine("字符缩放和间距设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"字符缩放和间距设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文本效果示例
        /// </summary>
        static void TextEffectsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加文本
                var range = document.Range();
                range.Text = "文本效果示例：普通文本、阴影文本、阳文文本、阴文文本。\n";

                // 设置阴影效果
                var shadowRange = document.Range(7, 11);
                shadowRange.Font.Shadow = true;


                Console.WriteLine("文本效果设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文本效果设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 多级标题和目录示例
        /// </summary>
        static void MultiLevelHeadingsAndTOCDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加多级标题
                var heading1 = document.Range();
                heading1.Text = "第一章 标题\n";
                heading1.Style = "标题 1";

                var heading2 = document.Range(document.Content.End - 1, document.Content.End - 1);
                heading2.Text = "1.1 节标题\n";
                heading2.Style = "标题 2";

                var heading3 = document.Range(document.Content.End - 1, document.Content.End - 1);
                heading3.Text = "1.1.1 小节标题\n";
                heading3.Style = "标题 3";

                // 添加正文内容
                var content = document.Range(document.Content.End - 1, document.Content.End - 1);
                content.Text = "这是正文内容。\n\n";

                // 插入目录
                var tocRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                tocRange.Text = "目录\n";
                tocRange.Font.Bold = true;
                tocRange.Font.Size = 16;

                // 添加目录项（模拟）
                var tocItem1 = document.Range(document.Content.End - 1, document.Content.End - 1);
                tocItem1.Text = "第一章 标题\t1\n";

                var tocItem2 = document.Range(document.Content.End - 1, document.Content.End - 1);
                tocItem2.Text = "1.1 节标题\t1\n";

                var tocItem3 = document.Range(document.Content.End - 1, document.Content.End - 1);
                tocItem3.Text = "1.1.1 小节标题\t1\n";

                Console.WriteLine("多级标题和目录设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"多级标题和目录设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例
        /// </summary>
        static void RealWorldApplicationDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                app.Visible = false; // 在实际应用示例中隐藏Word窗口

                var document = app.ActiveDocument;

                // 设置文档属性
                document.Title = "格式化文档示例";

                // 添加标题
                var title = document.Range();
                title.Text = "公司年度报告\n";
                title.Font.Name = "微软雅黑";
                title.Font.Size = 24;
                title.Font.Bold = true;
                title.Font.Color = WdColor.wdColorDarkBlue;
                title.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                title.ParagraphFormat.SpaceAfter = 24;

                // 添加副标题
                var subtitle = document.Range(document.Content.End - 1, document.Content.End - 1);
                subtitle.Text = "2025财年总结\n\n";
                subtitle.Font.Name = "微软雅黑";
                subtitle.Font.Size = 16;
                subtitle.Font.Bold = true;
                subtitle.Font.Color = WdColor.wdColorBlue;
                subtitle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                subtitle.ParagraphFormat.SpaceAfter = 18;

                // 添加章节标题
                var sectionTitle = document.Range(document.Content.End - 1, document.Content.End - 1);
                sectionTitle.Text = "财务概览\n";
                sectionTitle.Font.Name = "微软雅黑";
                sectionTitle.Font.Size = 14;
                sectionTitle.Font.Bold = true;
                sectionTitle.ParagraphFormat.SpaceAfter = 12;

                // 添加正文内容
                var content = document.Range(document.Content.End - 1, document.Content.End - 1);
                content.Text = "本年度公司实现了显著的财务增长，总收入达到1.2亿元，同比增长15%。净利润为3000万元，同比增长20%。\n\n";
                content.Font.Name = "宋体";
                content.Font.Size = 12;
                content.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                content.ParagraphFormat.FirstLineIndent = 21; // 首行缩进

                // 添加要点列表
                var points = document.Range(document.Content.End - 1, document.Content.End - 1);
                points.Text = "收入增长主要来源：\n• 产品线扩展\n• 市场份额提升\n• 新客户获取\n";
                points.Font.Name = "宋体";
                points.Font.Size = 12;

                // 添加表格数据
                var tableSection = document.Range(document.Content.End - 1, document.Content.End - 1);
                tableSection.Text = "\n关键财务指标：\n";
                tableSection.Font.Name = "微软雅黑";
                tableSection.Font.Size = 13;
                tableSection.Font.Bold = true;

                // 创建表格
                var tableRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                var table = document.Tables.Add(tableRange, 4, 3);
                table.Cell(1, 1).Range.Text = "指标";
                table.Cell(1, 2).Range.Text = "2024年";
                table.Cell(1, 3).Range.Text = "2025年";
                table.Cell(2, 1).Range.Text = "总收入(万元)";
                table.Cell(2, 2).Range.Text = "10,000";
                table.Cell(2, 3).Range.Text = "12,000";
                table.Cell(3, 1).Range.Text = "净利润(万元)";
                table.Cell(3, 2).Range.Text = "2,500";
                table.Cell(3, 3).Range.Text = "3,000";
                table.Cell(4, 1).Range.Text = "增长率";
                table.Cell(4, 2).Range.Text = "-";
                table.Cell(4, 3).Range.Text = "20%";

                // 格式化表格
                table.Borders.Enable = true;
                // 注意：由于库的限制，这里可能需要使用不同的方法来设置行对齐

                // 保存文档
                string filePath = Path.Combine(Path.GetTempPath(), "FormattedDocumentDemo.docx");
                document.SaveAs(filePath, WdSaveFormat.wdFormatDocumentDefault);

                Console.WriteLine($"格式化文档已创建: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                Console.WriteLine("使用TextFormattingManager辅助类进行完整操作:");

                // 创建文本格式化管理器实例
                var formattingManager = new TextFormattingManager();

                // 创建格式化文档
                var documentResult = formattingManager.CreateFormattedDocument();
                Console.WriteLine($"格式化文档创建结果:");
                Console.WriteLine($"  文档路径: {documentResult.DocumentPath}");
                Console.WriteLine($"  标题: {documentResult.Title}");
                Console.WriteLine($"  段落数: {documentResult.ParagraphCount}");

                // 应用字体格式
                var fontResult = formattingManager.ApplyFontFormatting();
                Console.WriteLine($"字体格式应用结果:");
                Console.WriteLine($"  原始文本: '{fontResult.OriginalText}'");
                Console.WriteLine($"  格式化文本: '{fontResult.FormattedText}'");

                // 应用段落格式
                var paragraphResult = formattingManager.ApplyParagraphFormatting();
                Console.WriteLine($"段落格式应用结果:");
                Console.WriteLine($"  段落对齐方式: {paragraphResult.Alignment}");
                Console.WriteLine($"  段前间距: {paragraphResult.SpaceBefore}");
                Console.WriteLine($"  段后间距: {paragraphResult.SpaceAfter}");

                // 应用高级格式化
                var advancedResult = formattingManager.ApplyAdvancedFormatting();
                Console.WriteLine($"高级格式化应用结果:");
                Console.WriteLine($"  高亮文本: '{advancedResult.HighlightedText}'");
                Console.WriteLine($"  下划线文本: '{advancedResult.UnderlinedText}'");
                Console.WriteLine($"  缩放文本: '{advancedResult.ScaledText}'");

                // 创建多级文档结构
                var structureResult = formattingManager.CreateStructuredDocument();
                Console.WriteLine($"多级文档结构创建结果:");
                Console.WriteLine($"  标题数量: {structureResult.HeadingCount}");
                Console.WriteLine($"  正文段落数: {structureResult.ParagraphCount}");
                Console.WriteLine($"  文档路径: {structureResult.DocumentPath}");

                Console.WriteLine("使用辅助类的完整示例操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例操作出错: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 文本格式化管理器辅助类
    /// </summary>
    public class TextFormattingManager
    {
        /// <summary>
        /// 文档创建结果类
        /// </summary>
        public class DocumentResult
        {
            /// <summary>
            /// 文档路径
            /// </summary>
            public string DocumentPath { get; set; }

            /// <summary>
            /// 标题
            /// </summary>
            public string Title { get; set; }

            /// <summary>
            /// 段落数
            /// </summary>
            public int ParagraphCount { get; set; }
        }

        /// <summary>
        /// 字体格式结果类
        /// </summary>
        public class FontFormattingResult
        {
            /// <summary>
            /// 原始文本
            /// </summary>
            public string OriginalText { get; set; }

            /// <summary>
            /// 格式化文本
            /// </summary>
            public string FormattedText { get; set; }
        }

        /// <summary>
        /// 段落格式结果类
        /// </summary>
        public class ParagraphFormattingResult
        {
            /// <summary>
            /// 段落对齐方式
            /// </summary>
            public WdParagraphAlignment Alignment { get; set; }

            /// <summary>
            /// 段前间距
            /// </summary>
            public float SpaceBefore { get; set; }

            /// <summary>
            /// 段后间距
            /// </summary>
            public float SpaceAfter { get; set; }
        }

        /// <summary>
        /// 高级格式化结果类
        /// </summary>
        public class AdvancedFormattingResult
        {
            /// <summary>
            /// 高亮文本
            /// </summary>
            public string HighlightedText { get; set; }

            /// <summary>
            /// 下划线文本
            /// </summary>
            public string UnderlinedText { get; set; }

            /// <summary>
            /// 缩放文本
            /// </summary>
            public string ScaledText { get; set; }
        }

        /// <summary>
        /// 结构化文档结果类
        /// </summary>
        public class StructuredDocumentResult
        {
            /// <summary>
            /// 标题数量
            /// </summary>
            public int HeadingCount { get; set; }

            /// <summary>
            /// 正文段落数
            /// </summary>
            public int ParagraphCount { get; set; }

            /// <summary>
            /// 文档路径
            /// </summary>
            public string DocumentPath { get; set; }
        }

        /// <summary>
        /// 创建格式化文档
        /// </summary>
        /// <returns>文档创建结果</returns>
        public DocumentResult CreateFormattedDocument()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;

            // 添加标题
            var title = document.Range();
            title.Text = "格式化文档示例\n";
            title.Font.Name = "微软雅黑";
            title.Font.Size = 18;
            title.Font.Bold = true;

            // 添加内容
            var content = document.Range(document.Content.End - 1, document.Content.End - 1);
            content.Text = "这是格式化文档的内容。\n\n包含多个段落用于演示格式化效果。\n\n这是最后一个段落。";

            // 保存文档
            string filePath = Path.Combine(Path.GetTempPath(), $"FormattedDocument_{Guid.NewGuid()}.docx");
            document.SaveAs(filePath, WdSaveFormat.wdFormatDocumentDefault);

            return new DocumentResult
            {
                DocumentPath = filePath,
                Title = "格式化文档示例",
                ParagraphCount = document.Paragraphs.Count
            };
        }

        /// <summary>
        /// 应用字体格式
        /// </summary>
        /// <returns>字体格式结果</returns>
        public FontFormattingResult ApplyFontFormatting()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;

            // 添加文本
            var range = document.Range();
            range.Text = "字体格式化示例文本";
            string originalText = range.Text;

            // 应用字体格式
            range.Font.Name = "楷体";
            range.Font.Size = 16;
            range.Font.Bold = true;
            range.Font.Italic = true;
            range.Font.Underline = WdUnderline.wdUnderlineSingle;
            range.Font.Color = WdColor.wdColorRed;

            return new FontFormattingResult
            {
                OriginalText = originalText,
                FormattedText = range.Text
            };
        }

        /// <summary>
        /// 应用段落格式
        /// </summary>
        /// <returns>段落格式结果</returns>
        public ParagraphFormattingResult ApplyParagraphFormatting()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;

            // 添加文本
            var range = document.Range();
            range.Text = "段落格式化示例文本。\n这是第二段文本。\n这是第三段文本。";

            // 应用段落格式
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            range.ParagraphFormat.SpaceBefore = 12;
            range.ParagraphFormat.SpaceAfter = 12;

            return new ParagraphFormattingResult
            {
                Alignment = range.ParagraphFormat.Alignment,
                SpaceBefore = range.ParagraphFormat.SpaceBefore,
                SpaceAfter = range.ParagraphFormat.SpaceAfter
            };
        }

        /// <summary>
        /// 应用高级格式化
        /// </summary>
        /// <returns>高级格式化结果</returns>
        public AdvancedFormattingResult ApplyAdvancedFormatting()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;

            // 添加文本
            var range = document.Range();
            range.Text = "高级格式化示例：高亮文本、下划线文本、缩放文本。";

            // 应用高亮
            var highlightRange = document.Range(8, 12);
            highlightRange.HighlightColorIndex = WdColorIndex.wdYellow;
            string highlightedText = highlightRange.Text;

            // 应用下划线
            var underlineRange = document.Range(13, 17);
            underlineRange.Font.Underline = WdUnderline.wdUnderlineDouble;
            string underlinedText = underlineRange.Text;

            // 应用字符缩放
            var scaleRange = document.Range(18, 22);
            // 注意：由于库的限制，这里可能需要使用不同的方法来设置字符缩放
            string scaledText = scaleRange.Text;

            return new AdvancedFormattingResult
            {
                HighlightedText = highlightedText,
                UnderlinedText = underlinedText,
                ScaledText = scaledText
            };
        }

        /// <summary>
        /// 创建结构化文档
        /// </summary>
        /// <returns>结构化文档结果</returns>
        public StructuredDocumentResult CreateStructuredDocument()
        {
            using var app = WordFactory.BlankWorkbook();
            var document = app.ActiveDocument;

            // 添加多级标题
            var heading1 = document.Range();
            heading1.Text = "第一章 简介\n";
            heading1.Style = "标题 1";

            var heading2 = document.Range(document.Content.End - 1, document.Content.End - 1);
            heading2.Text = "1.1 背景\n";
            heading2.Style = "标题 2";

            // 添加正文内容
            var content1 = document.Range(document.Content.End - 1, document.Content.End - 1);
            content1.Text = "这是第一章的正文内容。\n\n";

            var heading3 = document.Range(document.Content.End - 1, document.Content.End - 1);
            heading3.Text = "1.2 目标\n";
            heading3.Style = "标题 2";

            var content2 = document.Range(document.Content.End - 1, document.Content.End - 1);
            content2.Text = "这是第一章的另一个小节内容。\n";

            // 保存文档
            string filePath = Path.Combine(Path.GetTempPath(), $"StructuredDocument_{Guid.NewGuid()}.docx");
            document.SaveAs(filePath, WdSaveFormat.wdFormatDocumentDefault);

            return new StructuredDocumentResult
            {
                HeadingCount = 3,
                ParagraphCount = 2,
                DocumentPath = filePath
            };
        }
    }
}