using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindAndReplaceSample
{
    /// <summary>
    /// 文档清理助手类
    /// </summary>
    public class DocumentCleanupHelper
    {
        private readonly IWordDocument _document;
        private readonly FindAndReplaceHelper _findReplaceHelper;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public DocumentCleanupHelper(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _findReplaceHelper = new FindAndReplaceHelper(document);
        }

        /// <summary>
        /// 清理多余空格
        /// </summary>
        /// <returns>清理的空格数量</returns>
        public int CleanupExtraSpaces()
        {
            int totalReplacements = 0;

            try
            {
                // 替换多个空格为单个空格
                int replacements1 = 0;
                do
                {
                    replacements1 = _findReplaceHelper.ReplaceAll("  ", " "); // 两个空格替换为一个空格
                    totalReplacements += Math.Max(0, replacements1);
                } while (replacements1 > 0);

                // 清理行首空格
                int replacements2 = 0;
                do
                {
                    replacements2 = _findReplaceHelper.ReplaceAll("^p ", "^p"); // 段落标记后跟空格
                    totalReplacements += Math.Max(0, replacements2);
                } while (replacements2 > 0);

                Console.WriteLine("多余空格清理完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理多余空格时出错: {ex.Message}");
            }

            return totalReplacements;
        }

        /// <summary>
        /// 标准化称谓
        /// </summary>
        /// <returns>标准化的称谓数量</returns>
        public int StandardizeTitles()
        {
            int totalReplacements = 0;

            try
            {
                // 标准化公司名称和称谓
                var replacements = new Dictionary<string, string>
                {
                    {"Mr.", "先生"},
                    {"Mrs.", "夫人"},
                    {"Ms.", "女士"},
                    {"Dr.", "博士"},
                    {"Prof.", "教授"},
                    {"某某公司", "ABC有限公司"},
                    {"XYZ集团", "XYZ集团股份有限公司"},
                    {"DEF企业", "DEF企业发展有限公司"}
                };

                foreach (var pair in replacements)
                {
                    int count = _findReplaceHelper.ReplaceAll(pair.Key, pair.Value);
                    if (count > 0)
                    {
                        totalReplacements += (count == -1) ? 1 : count; // -1表示执行了替换但无法获取确切数量
                    }
                }

                Console.WriteLine("称谓标准化完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"标准化称谓时出错: {ex.Message}");
            }

            return totalReplacements;
        }

        /// <summary>
        /// 删除多余空行
        /// </summary>
        /// <returns>删除的空行数量</returns>
        public int RemoveExtraBlankLines()
        {
            int totalReplacements = 0;

            try
            {
                // 删除连续的空行（保留一个）
                int replacements1 = 0;
                do
                {
                    replacements1 = _findReplaceHelper.ReplaceAll("^p^p^p", "^p^p"); // 三个连续段落标记替换为两个
                    totalReplacements += Math.Max(0, replacements1);
                } while (replacements1 > 0);

                // 再次执行以处理更多连续空行
                int replacements2 = 0;
                do
                {
                    replacements2 = _findReplaceHelper.ReplaceAll("^p^p^p", "^p^p");
                    totalReplacements += Math.Max(0, replacements2);
                } while (replacements2 > 0);

                Console.WriteLine("空白行清理完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理空白行时出错: {ex.Message}");
            }

            return totalReplacements;
        }

        /// <summary>
        /// 标准化日期格式
        /// </summary>
        /// <returns>标准化的日期数量</returns>
        public int StandardizeDateFormats()
        {
            int totalReplacements = 0;

            try
            {
                // 查找 YYYY/MM/DD 格式并替换为 YYYY-MM-DD
                bool replaced1 = _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]{4})/([0-9]{2})/([0-9]{2})",
                    "\\1-\\2-\\3");

                if (replaced1)
                    totalReplacements++;

                // 查找 YYYY.MM.DD 格式并替换为 YYYY-MM-DD
                bool replaced2 = _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]{4})\\.([0-9]{2})\\.([0-9]{2})",
                    "\\1-\\2-\\3");

                if (replaced2)
                    totalReplacements++;

                Console.WriteLine("日期格式标准化完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"标准化日期格式时出错: {ex.Message}");
            }

            return totalReplacements;
        }

        /// <summary>
        /// 更新文档属性
        /// </summary>
        /// <param name="title">文档标题</param>
        /// <param name="author">文档作者</param>
        /// <param name="subject">文档主题</param>
        /// <param name="keywords">文档关键词</param>
        public void UpdateDocumentProperties(
            string title = "清理后的文档",
            string author = "文档清理工具",
            string subject = "已清理的文档",
            string keywords = "清理, 标准化, 自动化")
        {
            try
            {
                // 更新文档属性
                _document.Title = title;
                _document.Author = author;
                _document.Subject = subject;
                _document.Keywords = keywords;

                Console.WriteLine("文档属性更新完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"更新文档属性时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行完整的文档清理
        /// </summary>
        /// <param name="updateProperties">是否更新文档属性</param>
        /// <returns>清理报告</returns>
        public DocumentCleanupReport PerformFullCleanup(bool updateProperties = true)
        {
            var report = new DocumentCleanupReport();

            try
            {
                Console.WriteLine("开始文档清理...");

                // 1. 清理多余的空格
                report.SpacesCleaned = CleanupExtraSpaces();

                // 2. 标准化称谓
                report.TitlesStandardized = StandardizeTitles();

                // 3. 清理空白行
                report.BlankLinesRemoved = RemoveExtraBlankLines();

                // 4. 标准化日期格式
                report.DatesStandardized = StandardizeDateFormats();

                // 5. 更新文档属性
                if (updateProperties)
                {
                    UpdateDocumentProperties();
                }

                report.IsCompleted = true;
                Console.WriteLine("文档清理完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行完整文档清理时出错: {ex.Message}");
                report.IsCompleted = false;
                report.ErrorMessage = ex.Message;
            }

            return report;
        }

        /// <summary>
        /// 查找并高亮显示所有匹配项
        /// </summary>
        /// <param name="text">要查找的文本</param>
        /// <returns>匹配项数量</returns>
        public int HighlightAllMatches(string text)
        {
            int count = 0;

            try
            {
                var range = _document.Range();
                var find = range.Find;
                find.ClearFormatting();
                find.Text = text;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;

                while (find.Execute() && range.Start < range.End)
                {
                    // 高亮显示找到的文本
                    range.HighlightColorIndex = WdColorIndex.wdYellow;
                    count++;

                    // 移动到下一个位置
                    range = _document.Range(range.End, _document.Content.End);
                    find = range.Find;
                    find.Text = text;
                }

                Console.WriteLine($"已高亮显示 {count} 个匹配项");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"高亮显示匹配项时出错: {ex.Message}");
            }

            return count;
        }

        /// <summary>
        /// 移除所有高亮显示
        /// </summary>
        public void RemoveAllHighlights()
        {
            try
            {
                var range = _document.Range();
                range.HighlightColorIndex = WdColorIndex.wdNoHighlight;

                Console.WriteLine("已移除所有高亮显示");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"移除高亮显示时出错: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 文档清理报告类
    /// </summary>
    public class DocumentCleanupReport
    {
        /// <summary>
        /// 是否完成清理
        /// </summary>
        public bool IsCompleted { get; set; }

        /// <summary>
        /// 清理的空格数量
        /// </summary>
        public int SpacesCleaned { get; set; }

        /// <summary>
        /// 标准化的称谓数量
        /// </summary>
        public int TitlesStandardized { get; set; }

        /// <summary>
        /// 删除的空行数量
        /// </summary>
        public int BlankLinesRemoved { get; set; }

        /// <summary>
        /// 标准化的日期数量
        /// </summary>
        public int DatesStandardized { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成清理报告摘要
        /// </summary>
        /// <returns>报告摘要</returns>
        public string GenerateSummary()
        {
            if (!IsCompleted)
            {
                return $"清理未完成，错误信息: {ErrorMessage}";
            }

            return $"文档清理报告:\n" +
                   $"  清理多余空格: {SpacesCleaned} 处\n" +
                   $"  标准化称谓: {TitlesStandardized} 处\n" +
                   $"  删除空行: {BlankLinesRemoved} 处\n" +
                   $"  标准化日期: {DatesStandardized} 处\n" +
                   $"  清理完成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        }
    }
}