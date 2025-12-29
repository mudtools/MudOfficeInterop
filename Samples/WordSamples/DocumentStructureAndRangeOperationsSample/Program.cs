//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace DocumentStructureAndRangeOperationsSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 文档结构和范围操作示例");

            // 示例1: IWordDocument接口详解
            Console.WriteLine("\n=== 示例1: IWordDocument接口详解 ===");
            WordDocumentDemo();

            // 示例2: 文档基本属性和元数据
            Console.WriteLine("\n=== 示例2: 文档基本属性和元数据 ===");
            DocumentPropertiesDemo();

            // 示例3: 文档生命周期管理
            Console.WriteLine("\n=== 示例3: 文档生命周期管理 ===");
            DocumentLifecycleDemo();

            // 示例4: IWordRange接口详解
            Console.WriteLine("\n=== 示例4: IWordRange接口详解 ===");
            WordRangeDemo();

            // 示例5: 范围的选择和定义
            Console.WriteLine("\n=== 示例5: 范围的选择和定义 ===");
            RangeSelectionDemo();

            // 示例6: 文本内容操作
            Console.WriteLine("\n=== 示例6: 文本内容操作 ===");
            TextContentOperationsDemo();

            // 示例7: 范围复制和移动
            Console.WriteLine("\n=== 示例7: 范围复制和移动 ===");
            RangeCopyMoveDemo();

            // 示例8: 高级范围操作
            Console.WriteLine("\n=== 示例8: 高级范围操作 ===");
            AdvancedRangeOperationsDemo();

            // 示例9: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例9: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// IWordDocument接口详解示例
        /// </summary>
        static void WordDocumentDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 设置文档内容
                using var range = document.Range();
                range.Text = "文档标题\n\n这是文档正文内容。\n\n结束";

                // 获取文档基本信息
                string name = document.Name;
                string fullName = document.FullName;
                string path = document.Path;
                string title = document.Title;

                Console.WriteLine($"文档名称: {name}");
                Console.WriteLine($"完整路径: {fullName}");
                Console.WriteLine($"文档路径: {path}");
                Console.WriteLine($"文档标题: {title}");

                // 检查文档状态
                bool? saved = document.Saved;
                bool? routed = document.Routed;

                Console.WriteLine($"是否已保存: {saved}");
                Console.WriteLine($"是否已发送路由: {routed}");

                // 设置文档属性
                document.Title = "新标题";
                Console.WriteLine($"更新后的文档标题: {document.Title}");

                Console.WriteLine("IWordDocument接口操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"IWordDocument接口操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文档基本属性和元数据示例
        /// </summary>
        static void DocumentPropertiesDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 添加示例内容
                using var range = document.Range();
                range.Text = "示例文档内容\n\n创建时间: " + DateTime.Now.ToString();

                // 文档基本信息
                Console.WriteLine($"文档名称: {document.Name}");
                Console.WriteLine($"完整路径: {document.FullName}");
                Console.WriteLine($"文档路径: {document.Path}");
                Console.WriteLine($"文档标题: {document.Title}");

                // 文档状态
                Console.WriteLine($"是否已保存: {document.Saved}");
                Console.WriteLine($"是否已发送路由: {document.Routed}");
                Console.WriteLine($"是否为主控文档: {document.IsMasterDocument}");

                // 字体嵌入设置
                document.EmbedTrueTypeFonts = true;
                document.SaveSubsetFonts = true;
                Console.WriteLine("已设置字体嵌入选项");

                // 保护设置
                document.ReadOnlyRecommended = true;
                Console.WriteLine("已设置只读推荐");

                Console.WriteLine("文档基本属性和元数据操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文档基本属性和元数据操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文档生命周期管理示例
        /// </summary>
        static void DocumentLifecycleDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 编辑文档内容
                using var range = document.Range();
                range.Text = "文档生命周期管理示例\n\n内容创建时间: " + DateTime.Now.ToString();

                // 保存文档
                string filePath = Path.Combine(Path.GetTempPath(), "LifecycleDemo.docx");
                document.SaveAs(filePath);
                Console.WriteLine($"文档已保存到: {filePath}");
                Console.WriteLine($"保存状态: {document.Saved}");

                Console.WriteLine("文档生命周期管理操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文档生命周期管理操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// IWordRange接口详解示例
        /// </summary>
        static void WordRangeDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 获取整个文档的范围
                using var range = document.Range();
                range.Text = "第一段文本。\n第二段文本。\n第三段文本。";

                // 创建指定位置的范围
                var specificRange = document.Range(0, 5);
                Console.WriteLine($"指定范围文本: {specificRange.Text}");

                // 获取范围的副本
                var duplicateRange = range.Duplicate;
                Console.WriteLine($"范围副本文本长度: {duplicateRange.Text.Length}");

                Console.WriteLine("IWordRange接口操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"IWordRange接口操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 范围的选择和定义示例
        /// </summary>
        static void RangeSelectionDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 获取整个文档范围并填充示例文本
                using var range = document.Range();
                range.Text = "这是第一段文本。\n这是第二段文本。\n这是第三段文本。";

                // 重新定义范围
                range.Start = 0;
                range.End = 5;
                Console.WriteLine($"范围文本: {range.Text}");

                // 移动范围
                range.Start = 6;
                range.End = 12;
                Console.WriteLine($"新范围文本: {range.Text}");

                Console.WriteLine("范围的选择和定义操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"范围的选择和定义操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文本内容操作示例
        /// </summary>
        static void TextContentOperationsDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;
                var range = document.Range();

                // 设置文本
                range.Text = "Hello World!";
                Console.WriteLine($"初始文本: {range.Text}");

                // 插入文本
                range.InsertBefore("前缀文本 ");
                range.InsertAfter(" 后缀文本");
                Console.WriteLine($"插入后文本: {range.Text}");

                // 删除文本
                var deleteRange = document.Range(0, 5);
                deleteRange.Delete();
                Console.WriteLine($"删除后文本: {document.Range().Text}");

                // 替换文本
                range = document.Range();
                range.Text = range.Text.Replace("Hello", "Hi");
                Console.WriteLine($"替换后文本: {range.Text}");

                Console.WriteLine("文本内容操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文本内容操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 范围复制和移动示例
        /// </summary>
        static void RangeCopyMoveDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 添加内容
                using var range1 = document.Range();
                range1.Text = "原始内容\n";

                using var range2 = document.Range();
                range2.InsertAfter("另一部分内容\n");

                Console.WriteLine($"复制前文档内容:\n{document.Range().Text}");

                // 复制内容
                using var sourceRange = document.Range(0, 4);
                using var targetRange = document.Range(document.Content.End, document.Content.End);
                sourceRange.Copy();
                targetRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                targetRange.Paste();

                Console.WriteLine($"复制后文档内容:\n{document.Range().Text}");

                // 移动内容
                using var moveSource = document.Range(5, 9);
                using var moveTarget = document.Range(document.Content.End, document.Content.End);
                moveSource.Cut();
                moveTarget.Collapse(WdCollapseDirection.wdCollapseEnd);
                moveTarget.Paste();

                Console.WriteLine($"移动后文档内容:\n{document.Range().Text}");

                Console.WriteLine("范围复制和移动操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"范围复制和移动操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 高级范围操作示例
        /// </summary>
        static void AdvancedRangeOperationsDemo()
        {
            try
            {
                using var app = WordFactory.BlankDocument();
                using var document = app.ActiveDocument;

                // 创建复杂文档结构
                using var range = document.Range();
                range.Text = "第一段文本内容。\n\n第二段文本内容，包含更多信息。\n\n第三段文本内容，这是最后一段。";

                // 查找和替换操作
                using var findRange = document.Range();
                using var find = findRange.Find;
                find.ClearFormatting();
                find.Text = "文本";
                find.Replacement.Text = "内容";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindContinue;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;
                find.Execute(replace: WdReplace.wdReplaceAll);

                Console.WriteLine($"查找替换后文档内容:\n{document.Range().Text}");

                // 获取段落数量
                int paragraphCount = document.Paragraphs.Count;
                Console.WriteLine($"文档段落数量: {paragraphCount}");

                // 获取字符数
                int characterCount = document.Characters.Count;
                Console.WriteLine($"文档字符数: {characterCount}");

                Console.WriteLine("高级范围操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"高级范围操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                Console.WriteLine("使用DocumentStructureManager辅助类进行完整操作:");

                // 创建文档结构管理器实例
                var structureManager = new DocumentStructureManager();

                // 创建并分析文档
                var analysisResult = structureManager.CreateAndAnalyzeDocument();
                Console.WriteLine($"文档分析结果:");
                Console.WriteLine($"  文档名称: {analysisResult.DocumentName}");
                Console.WriteLine($"  段落数量: {analysisResult.ParagraphCount}");
                Console.WriteLine($"  字符数量: {analysisResult.CharacterCount}");
                Console.WriteLine($"  单词数量: {analysisResult.WordCount}");

                // 执行范围操作
                var rangeOperationsResult = structureManager.PerformRangeOperations();
                Console.WriteLine($"范围操作结果:");
                Console.WriteLine($"  原始文本: {rangeOperationsResult.OriginalText}");
                Console.WriteLine($"  修改后文本: {rangeOperationsResult.ModifiedText}");

                // 执行查找替换操作
                var findReplaceResult = structureManager.PerformFindReplaceOperations();
                Console.WriteLine($"查找替换结果:");
                Console.WriteLine($"  操作前文本: {findReplaceResult.BeforeText}");
                Console.WriteLine($"  操作后文本: {findReplaceResult.AfterText}");
                Console.WriteLine($"  替换次数: {findReplaceResult.ReplacementCount}");

                Console.WriteLine("使用辅助类的完整示例操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例操作出错: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 文档结构管理器辅助类
    /// </summary>
    public class DocumentStructureManager
    {
        /// <summary>
        /// 文档分析结果类
        /// </summary>
        public class DocumentAnalysisResult
        {
            /// <summary>
            /// 文档名称
            /// </summary>
            public string DocumentName { get; set; }

            /// <summary>
            /// 段落数量
            /// </summary>
            public int ParagraphCount { get; set; }

            /// <summary>
            /// 字符数量
            /// </summary>
            public int CharacterCount { get; set; }

            /// <summary>
            /// 单词数量
            /// </summary>
            public int WordCount { get; set; }
        }

        /// <summary>
        /// 范围操作结果类
        /// </summary>
        public class RangeOperationsResult
        {
            /// <summary>
            /// 原始文本
            /// </summary>
            public string OriginalText { get; set; }

            /// <summary>
            /// 修改后文本
            /// </summary>
            public string ModifiedText { get; set; }
        }

        /// <summary>
        /// 查找替换操作结果类
        /// </summary>
        public class FindReplaceResult
        {
            /// <summary>
            /// 操作前文本
            /// </summary>
            public string BeforeText { get; set; }

            /// <summary>
            /// 操作后文本
            /// </summary>
            public string AfterText { get; set; }

            /// <summary>
            /// 替换次数
            /// </summary>
            public int ReplacementCount { get; set; }
        }

        /// <summary>
        /// 创建并分析文档
        /// </summary>
        /// <returns>文档分析结果</returns>
        public DocumentAnalysisResult CreateAndAnalyzeDocument()
        {
            using var app = WordFactory.BlankDocument();
            using var document = app.ActiveDocument;

            // 添加示例内容
            using var range = document.Range();
            range.Text = "这是第一段文本内容，用于分析。\n\n" +
                        "这是第二段文本内容，包含更多信息用于详细分析。\n\n" +
                        "这是第三段文本内容，作为最后一个段落用于完整分析。";

            // 返回分析结果
            return new DocumentAnalysisResult
            {
                DocumentName = document.Name,
                ParagraphCount = document.Paragraphs.Count,
                CharacterCount = document.Characters.Count,
                WordCount = document.Words.Count
            };
        }

        /// <summary>
        /// 执行范围操作
        /// </summary>
        /// <returns>范围操作结果</returns>
        public RangeOperationsResult PerformRangeOperations()
        {
            using var app = WordFactory.BlankDocument();
            using var document = app.ActiveDocument;

            // 添加初始内容
            using var range = document.Range();
            range.Text = "初始文本内容用于范围操作演示。";

            string originalText = range.Text;

            // 执行范围操作
            range.InsertBefore("前缀: ");
            range.InsertAfter(" 后缀");
            range.Font.Bold = true;

            return new RangeOperationsResult
            {
                OriginalText = originalText,
                ModifiedText = document.Range().Text
            };
        }

        /// <summary>
        /// 执行查找替换操作
        /// </summary>
        /// <returns>查找替换操作结果</returns>
        public FindReplaceResult PerformFindReplaceOperations()
        {
            using var app = WordFactory.BlankDocument();
            using var document = app.ActiveDocument;

            // 添加初始内容
            using var range = document.Range();
            range.Text = "查找替换操作演示。查找文本并替换文本。多次出现文本需要替换。";

            string beforeText = range.Text;
            int replacementCount = 0;

            // 执行查找替换操作
            using var findRange = document.Range();
            using var find = findRange.Find;
            find.ClearFormatting();
            find.Text = "文本";
            find.Replacement.Text = "内容";
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindContinue;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;
            replacementCount = find.Execute(replace: WdReplace.wdReplaceAll) ? 1 : 0; // 简化处理

            return new FindReplaceResult
            {
                BeforeText = beforeText,
                AfterText = document.Range().Text,
                ReplacementCount = replacementCount
            };
        }
    }
}