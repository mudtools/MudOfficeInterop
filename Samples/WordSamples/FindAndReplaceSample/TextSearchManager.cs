//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace FindAndReplaceSample
{
    /// <summary>
    /// 文本搜索管理器类
    /// </summary>
    public class TextSearchManager
    {
        private readonly IWordDocument _document;
        private readonly FindAndReplaceHelper _findReplaceHelper;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public TextSearchManager(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _findReplaceHelper = new FindAndReplaceHelper(document);
        }

        /// <summary>
        /// 查找结果信息类
        /// </summary>
        public class SearchResult
        {
            /// <summary>
            /// 起始位置
            /// </summary>
            public int StartPosition { get; set; }

            /// <summary>
            /// 结束位置
            /// </summary>
            public int EndPosition { get; set; }

            /// <summary>
            /// 匹配文本
            /// </summary>
            public string MatchedText { get; set; }

            /// <summary>
            /// 匹配文本长度
            /// </summary>
            public int Length => EndPosition - StartPosition;
        }

        /// <summary>
        /// 查找所有匹配项
        /// </summary>
        /// <param name="text">要查找的文本</param>
        /// <param name="caseSensitive">是否区分大小写</param>
        /// <param name="wholeWord">是否全字匹配</param>
        /// <returns>匹配结果列表</returns>
        public List<SearchResult> FindAllOccurrences(string text, bool caseSensitive = false, bool wholeWord = false)
        {
            var results = new List<SearchResult>();

            try
            {
                var range = _document.Range();
                var find = range.Find;
                find.ClearFormatting();
                find.Text = text;
                find.MatchCase = caseSensitive;
                find.MatchWholeWord = wholeWord;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;

                while (find.Execute() && range.Start < range.End)
                {
                    var result = new SearchResult
                    {
                        StartPosition = range.Start,
                        EndPosition = range.End,
                        MatchedText = range.Text
                    };

                    results.Add(result);

                    // 移动到下一个位置
                    range = _document.Range(range.End, _document.Content.End);
                    find = range.Find;
                    find.Text = text;
                    find.MatchCase = caseSensitive;
                    find.MatchWholeWord = wholeWord;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找所有匹配项时出错: {ex.Message}");
            }

            return results;
        }

        /// <summary>
        /// 查找并替换所有匹配项，同时返回替换结果
        /// </summary>
        /// <param name="findText">要查找的文本</param>
        /// <param name="replaceWith">替换文本</param>
        /// <param name="caseSensitive">是否区分大小写</param>
        /// <param name="wholeWord">是否全字匹配</param>
        /// <returns>替换结果</returns>
        public TextReplaceResult ReplaceAllWithReport(string findText, string replaceWith, bool caseSensitive = false, bool wholeWord = false)
        {
            var result = new TextReplaceResult
            {
                FindText = findText,
                ReplaceWith = replaceWith,
                CaseSensitive = caseSensitive,
                WholeWord = wholeWord,
                MatchesFound = new List<SearchResult>()
            };

            try
            {
                // 先查找所有匹配项
                result.MatchesFound = FindAllOccurrences(findText, caseSensitive, wholeWord);
                result.OccurrencesFound = result.MatchesFound.Count;

                // 执行替换
                using var find = _document.Range().Find;
                find.ClearFormatting();
                find.Text = findText;
                find.Replacement.ClearFormatting();
                find.Replacement.Text = replaceWith;
                find.MatchCase = caseSensitive;
                find.MatchWholeWord = wholeWord;

                find.Execute(
                    findText: findText,
                    replaceWith: replaceWith,
                    replace: WdReplace.wdReplaceAll
                );

                result.ReplacedSuccessfully = true;
                result.OccurrencesReplaced = result.OccurrencesFound;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找并替换时出错: {ex.Message}");
                result.ReplacedSuccessfully = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 使用通配符查找所有匹配项
        /// </summary>
        /// <param name="pattern">通配符模式</param>
        /// <returns>匹配结果列表</returns>
        public List<SearchResult> FindAllWithWildcards(string pattern)
        {
            var results = new List<SearchResult>();

            try
            {
                var range = _document.Range();
                var find = range.Find;
                find.ClearFormatting();
                find.Text = pattern;
                find.MatchWildcards = true;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;

                while (find.Execute() && range.Start < range.End)
                {
                    var result = new SearchResult
                    {
                        StartPosition = range.Start,
                        EndPosition = range.End,
                        MatchedText = range.Text
                    };

                    results.Add(result);

                    // 移动到下一个位置
                    range = _document.Range(range.End, _document.Content.End);
                    find = range.Find;
                    find.Text = pattern;
                    find.MatchWildcards = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用通配符查找所有匹配项时出错: {ex.Message}");
            }

            return results;
        }

        /// <summary>
        /// 基于格式查找所有匹配项
        /// </summary>
        /// <param name="bold">是否粗体</param>
        /// <param name="italic">是否斜体</param>
        /// <param name="underline">下划线类型</param>
        /// <returns>匹配结果列表</returns>
        public List<SearchResult> FindAllByFormat(bool? bold = null, bool? italic = null, WdUnderline? underline = null)
        {
            var results = new List<SearchResult>();

            try
            {
                var range = _document.Range();
                var find = range.Find;
                find.ClearFormatting();

                if (bold.HasValue)
                    find.Font.Bold = bold.Value;

                if (italic.HasValue)
                    find.Font.Italic = italic.Value;

                if (underline.HasValue)
                    find.Font.Underline = underline.Value;

                find.Text = ""; // 只基于格式查找
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;

                while (find.Execute() && range.Start < range.End)
                {
                    var result = new SearchResult
                    {
                        StartPosition = range.Start,
                        EndPosition = range.End,
                        MatchedText = range.Text
                    };

                    results.Add(result);

                    // 移动到下一个位置
                    range = _document.Range(range.End, _document.Content.End);
                    find = range.Find;

                    if (bold.HasValue)
                        find.Font.Bold = bold.Value;

                    if (italic.HasValue)
                        find.Font.Italic = italic.Value;

                    if (underline.HasValue)
                        find.Font.Underline = underline.Value;

                    find.Text = "";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"基于格式查找所有匹配项时出错: {ex.Message}");
            }

            return results;
        }

        /// <summary>
        /// 查找电话号码
        /// </summary>
        /// <returns>电话号码列表</returns>
        public List<SearchResult> FindPhoneNumbers()
        {
            return FindAllWithWildcards("[0-9]{3}-[0-9]{4}-[0-9]{4}");
        }

        /// <summary>
        /// 查找邮箱地址
        /// </summary>
        /// <returns>邮箱地址列表</returns>
        public List<SearchResult> FindEmailAddresses()
        {
            return FindAllWithWildcards("[a-zA-Z0-9]*@[a-zA-Z0-9]*\\.[a-zA-Z]*");
        }

        /// <summary>
        /// 查找日期格式
        /// </summary>
        /// <returns>日期列表</returns>
        public List<SearchResult> FindDates()
        {
            return FindAllWithWildcards("[0-9]{4}-[0-9]{2}-[0-9]{2}");
        }

        /// <summary>
        /// 创建查找范围
        /// </summary>
        /// <param name="start">起始位置</param>
        /// <param name="end">结束位置</param>
        /// <returns>查找范围</returns>
        public IWordRange CreateSearchRange(int start, int end)
        {
            try
            {
                return _document.Range(start, end);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建查找范围时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取文档统计信息
        /// </summary>
        /// <returns>文档统计信息</returns>
        public DocumentStatistics GetDocumentStatistics()
        {
            var stats = new DocumentStatistics();

            try
            {
                stats.TotalCharacters = _document.Characters.Count;
                stats.TotalWords = _document.Words.Count;
                stats.TotalParagraphs = _document.Paragraphs.Count;
                stats.TotalPages = _document.Range().Paragraphs.Count / 50; // 粗略估算
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取文档统计信息时出错: {ex.Message}");
            }

            return stats;
        }

        /// <summary>
        /// 高级文本搜索选项
        /// </summary>
        public class AdvancedSearchOptions
        {
            /// <summary>
            /// 是否区分大小写
            /// </summary>
            public bool MatchCase { get; set; }

            /// <summary>
            /// 是否全字匹配
            /// </summary>
            public bool MatchWholeWord { get; set; }

            /// <summary>
            /// 是否使用通配符
            /// </summary>
            public bool MatchWildcards { get; set; }

            /// <summary>
            /// 是否模糊匹配
            /// </summary>
            public bool MatchFuzzy { get; set; }

            /// <summary>
            /// 是否匹配 Kashida 字符
            /// </summary>
            public bool MatchKashida { get; set; }
        }

        /// <summary>
        /// 使用高级选项进行搜索
        /// </summary>
        /// <param name="text">查找文本</param>
        /// <param name="options">高级搜索选项</param>
        /// <returns>是否找到</returns>
        public bool AdvancedSearch(string text, AdvancedSearchOptions options)
        {
            try
            {
                using var find = _document.Range().Find;
                find.ClearFormatting();
                find.Text = text;
                find.MatchCase = options.MatchCase;
                find.MatchWholeWord = options.MatchWholeWord;
                find.MatchWildcards = options.MatchWildcards;
                find.MatchFuzzy = options.MatchFuzzy;
                find.MatchKashida = options.MatchKashida;

                bool found = find.Execute();
                return found;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"高级搜索时出错: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// 文本替换结果类
    /// </summary>
    public class TextReplaceResult
    {
        /// <summary>
        /// 查找的文本
        /// </summary>
        public string FindText { get; set; }

        /// <summary>
        /// 替换的文本
        /// </summary>
        public string ReplaceWith { get; set; }

        /// <summary>
        /// 是否区分大小写
        /// </summary>
        public bool CaseSensitive { get; set; }

        /// <summary>
        /// 是否全字匹配
        /// </summary>
        public bool WholeWord { get; set; }

        /// <summary>
        /// 找到的匹配项数量
        /// </summary>
        public int OccurrencesFound { get; set; }

        /// <summary>
        /// 替换的匹配项数量
        /// </summary>
        public int OccurrencesReplaced { get; set; }

        /// <summary>
        /// 是否替换成功
        /// </summary>
        public bool ReplacedSuccessfully { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 找到的匹配项列表
        /// </summary>
        public List<TextSearchManager.SearchResult> MatchesFound { get; set; }

        /// <summary>
        /// 生成替换结果报告
        /// </summary>
        /// <returns>结果报告</returns>
        public string GenerateReport()
        {
            if (!ReplacedSuccessfully)
            {
                return $"替换失败: {ErrorMessage}";
            }

            return $"文本替换报告:\n" +
                   $"  查找文本: '{FindText}'\n" +
                   $"  替换为: '{ReplaceWith}'\n" +
                   $"  区分大小写: {CaseSensitive}\n" +
                   $"  全字匹配: {WholeWord}\n" +
                   $"  找到匹配项: {OccurrencesFound} 个\n" +
                   $"  替换完成: {OccurrencesReplaced} 个\n" +
                   $"  替换时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        }
    }

    /// <summary>
    /// 文档统计信息类
    /// </summary>
    public class DocumentStatistics
    {
        /// <summary>
        /// 总字符数
        /// </summary>
        public int TotalCharacters { get; set; }

        /// <summary>
        /// 总单词数
        /// </summary>
        public int TotalWords { get; set; }

        /// <summary>
        /// 总段落数
        /// </summary>
        public int TotalParagraphs { get; set; }

        /// <summary>
        /// 总页数（估算）
        /// </summary>
        public int TotalPages { get; set; }

        /// <summary>
        /// 生成统计报告
        /// </summary>
        /// <returns>统计报告</returns>
        public string GenerateReport()
        {
            return $"文档统计信息:\n" +
                   $"  总字符数: {TotalCharacters}\n" +
                   $"  总单词数: {TotalWords}\n" +
                   $"  总段落数: {TotalParagraphs}\n" +
                   $"  估算页数: {TotalPages}";
        }
    }
}