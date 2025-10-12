using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindAndReplaceSample
{
    /// <summary>
    /// 模式替换助手类
    /// </summary>
    public class PatternReplaceHelper
    {
        private readonly IWordDocument _document;
        private readonly FindAndReplaceHelper _findReplaceHelper;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public PatternReplaceHelper(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _findReplaceHelper = new FindAndReplaceHelper(document);
        }

        /// <summary>
        /// 电话号码格式标准化
        /// </summary>
        /// <returns>是否执行成功</returns>
        public bool StandardizePhoneNumbers()
        {
            try
            {
                // 将各种电话号码格式统一为 138-1234-5678 格式
                bool success = true;

                // 处理 13812345678 格式
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]{3})([0-9]{4})([0-9]{4})",
                    "\\1-\\2-\\3");

                // 处理 138 1234 5678 格式
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]{3}) ([0-9]{4}) ([0-9]{4})",
                    "\\1-\\2-\\3");

                // 处理 138.1234.5678 格式
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]{3})\\.([0-9]{4})\\.([0-9]{4})",
                    "\\1-\\2-\\3");

                Console.WriteLine("电话号码格式标准化完成");
                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"电话号码格式标准化时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 邮箱地址格式标准化
        /// </summary>
        /// <returns>是否执行成功</returns>
        public bool StandardizeEmailAddresses()
        {
            try
            {
                // 将邮箱地址转换为小写并标准化格式
                bool success = true;

                // 处理常见的邮箱格式
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "([a-zA-Z0-9]*)@([a-zA-Z0-9]*)\\.([a-zA-Z]*)",
                    "\\1@\\2.\\3");

                Console.WriteLine("邮箱地址格式标准化完成");
                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"邮箱地址格式标准化时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 日期格式标准化
        /// </summary>
        /// <returns>是否执行成功</returns>
        public bool StandardizeDateFormats()
        {
            try
            {
                bool success = true;

                // 处理 YYYY/MM/DD 格式
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]{4})/([0-9]{2})/([0-9]{2})",
                    "\\1-\\2-\\3");

                // 处理 YYYY.MM.DD 格式
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]{4})\\.([0-9]{2})\\.([0-9]{2})",
                    "\\1-\\2-\\3");

                // 处理 DD-MM-YYYY 格式
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]{2})-([0-9]{2})-([0-9]{4})",
                    "\\3-\\2-\\1");

                // 处理 MM/DD/YYYY 格式
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]{2})/([0-9]{2})/([0-9]{4})",
                    "\\3-\\1-\\2");

                Console.WriteLine("日期格式标准化完成");
                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"日期格式标准化时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 数字格式标准化
        /// </summary>
        /// <returns>是否执行成功</returns>
        public bool StandardizeNumberFormats()
        {
            try
            {
                bool success = true;

                // 处理千位分隔符（将逗号替换为标准格式）
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "([0-9]*),([0-9]{3})",
                    "\\1,\\2");

                Console.WriteLine("数字格式标准化完成");
                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"数字格式标准化时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// URL格式标准化
        /// </summary>
        /// <returns>是否执行成功</returns>
        public bool StandardizeUrls()
        {
            try
            {
                bool success = true;

                // 处理常见的URL格式
                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "http://([a-zA-Z0-9\\.\\-]*)",
                    "http://\\1");

                success &= _findReplaceHelper.ReplaceWithWildcards(
                    "https://([a-zA-Z0-9\\.\\-]*)",
                    "https://\\1");

                Console.WriteLine("URL格式标准化完成");
                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"URL格式标准化时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 标点符号标准化
        /// </summary>
        /// <returns>是否执行成功</returns>
        public bool StandardizePunctuation()
        {
            try
            {
                bool success = true;

                // 处理多个连续的句号
                success &= _findReplaceHelper.ReplaceAll("。。。", "……");

                // 处理多个连续的感叹号
                success &= _findReplaceHelper.ReplaceAll("！！！", "！");

                // 处理多个连续的问号
                success &= _findReplaceHelper.ReplaceAll("？？？", "？");

                Console.WriteLine("标点符号标准化完成");
                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"标点符号标准化时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 英文单词大小写标准化
        /// </summary>
        /// <returns>是否执行成功</returns>
        public bool StandardizeEnglishCapitalization()
        {
            try
            {
                bool success = true;

                // 标准化常见的英文单词大小写
                var capitalizations = new Dictionary<string, string>
                {
                    {"internet", "Internet"},
                    {"email", "Email"},
                    {"website", "Website"},
                    {"online", "Online"},
                    {"url", "URL"},
                    {"html", "HTML"},
                    {"css", "CSS"},
                    {"javascript", "JavaScript"},
                    {"api", "API"}
                };

                foreach (var pair in capitalizations)
                {
                    success &= _findReplaceHelper.ReplaceAll(pair.Key, pair.Value);
                }

                Console.WriteLine("英文单词大小写标准化完成");
                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"英文单词大小写标准化时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 执行所有模式替换
        /// </summary>
        /// <returns>标准化报告</returns>
        public PatternStandardizationReport PerformAllStandardizations()
        {
            var report = new PatternStandardizationReport();

            try
            {
                Console.WriteLine("开始执行所有模式标准化...");

                report.PhoneNumbersStandardized = StandardizePhoneNumbers();
                report.EmailAddressesStandardized = StandardizeEmailAddresses();
                report.DateFormatsStandardized = StandardizeDateFormats();
                report.NumberFormatsStandardized = StandardizeNumberFormats();
                report.UrlsStandardized = StandardizeUrls();
                report.PunctuationStandardized = StandardizePunctuation();
                report.EnglishCapitalizationStandardized = StandardizeEnglishCapitalization();

                report.IsCompleted = true;
                Console.WriteLine("所有模式标准化完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行所有模式标准化时出错: {ex.Message}");
                report.IsCompleted = false;
                report.ErrorMessage = ex.Message;
            }

            return report;
        }

        /// <summary>
        /// 自定义模式替换
        /// </summary>
        /// <param name="patterns">模式替换字典</param>
        /// <returns>替换结果</returns>
        public CustomPatternReplaceResult ReplaceCustomPatterns(Dictionary<string, string> patterns)
        {
            var result = new CustomPatternReplaceResult
            {
                Patterns = patterns,
                SuccessfulReplacements = new List<string>(),
                FailedReplacements = new List<string>()
            };

            try
            {
                foreach (var pattern in patterns)
                {
                    bool success = _findReplaceHelper.ReplaceWithWildcards(pattern.Key, pattern.Value);
                    if (success)
                    {
                        result.SuccessfulReplacements.Add(pattern.Key);
                    }
                    else
                    {
                        result.FailedReplacements.Add(pattern.Key);
                    }
                }

                result.IsCompleted = true;
                Console.WriteLine("自定义模式替换完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自定义模式替换时出错: {ex.Message}");
                result.IsCompleted = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 创建自定义模式替换
        /// </summary>
        /// <param name="findPattern">查找模式</param>
        /// <param name="replacePattern">替换模式</param>
        /// <returns>是否执行成功</returns>
        public bool CreateCustomPatternReplace(string findPattern, string replacePattern)
        {
            try
            {
                bool success = _findReplaceHelper.ReplaceWithWildcards(findPattern, replacePattern);
                Console.WriteLine($"自定义模式替换完成: {findPattern} -> {replacePattern}");
                return success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建自定义模式替换时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 批量模式替换
        /// </summary>
        /// <param name="patternReplacements">模式替换列表</param>
        /// <returns>批量替换结果</returns>
        public BatchPatternReplaceResult BatchReplace(List<PatternReplacement> patternReplacements)
        {
            var result = new BatchPatternReplaceResult
            {
                Replacements = patternReplacements,
                Results = new List<PatternReplaceResult>()
            };

            try
            {
                foreach (var replacement in patternReplacements)
                {
                    bool success = _findReplaceHelper.ReplaceWithWildcards(
                        replacement.FindPattern, 
                        replacement.ReplacePattern);

                    var replaceResult = new PatternReplaceResult
                    {
                        FindPattern = replacement.FindPattern,
                        ReplacePattern = replacement.ReplacePattern,
                        IsSuccessful = success
                    };

                    result.Results.Add(replaceResult);
                }

                result.IsCompleted = true;
                Console.WriteLine("批量模式替换完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量模式替换时出错: {ex.Message}");
                result.IsCompleted = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }
    }

    /// <summary>
    /// 模式替换结果类
    /// </summary>
    public class PatternReplaceResult
    {
        /// <summary>
        /// 查找模式
        /// </summary>
        public string FindPattern { get; set; }

        /// <summary>
        /// 替换模式
        /// </summary>
        public string ReplacePattern { get; set; }

        /// <summary>
        /// 是否执行成功
        /// </summary>
        public bool IsSuccessful { get; set; }
    }

    /// <summary>
    /// 模式替换类
    /// </summary>
    public class PatternReplacement
    {
        /// <summary>
        /// 查找模式
        /// </summary>
        public string FindPattern { get; set; }

        /// <summary>
        /// 替换模式
        /// </summary>
        public string ReplacePattern { get; set; }

        /// <summary>
        /// 描述
        /// </summary>
        public string Description { get; set; }
    }

    /// <summary>
    /// 模式标准化报告类
    /// </summary>
    public class PatternStandardizationReport
    {
        /// <summary>
        /// 是否完成标准化
        /// </summary>
        public bool IsCompleted { get; set; }

        /// <summary>
        /// 电话号码标准化是否成功
        /// </summary>
        public bool PhoneNumbersStandardized { get; set; }

        /// <summary>
        /// 邮箱地址标准化是否成功
        /// </summary>
        public bool EmailAddressesStandardized { get; set; }

        /// <summary>
        /// 日期格式标准化是否成功
        /// </summary>
        public bool DateFormatsStandardized { get; set; }

        /// <summary>
        /// 数字格式标准化是否成功
        /// </summary>
        public bool NumberFormatsStandardized { get; set; }

        /// <summary>
        /// URL标准化是否成功
        /// </summary>
        public bool UrlsStandardized { get; set; }

        /// <summary>
        /// 标点符号标准化是否成功
        /// </summary>
        public bool PunctuationStandardized { get; set; }

        /// <summary>
        /// 英文单词大小写标准化是否成功
        /// </summary>
        public bool EnglishCapitalizationStandardized { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成标准化报告
        /// </summary>
        /// <returns>标准化报告</returns>
        public string GenerateReport()
        {
            if (!IsCompleted)
            {
                return $"标准化未完成，错误信息: {ErrorMessage}";
            }

            return $"模式标准化报告:\n" +
                   $"  电话号码标准化: {PhoneNumbersStandardized}\n" +
                   $"  邮箱地址标准化: {EmailAddressesStandardized}\n" +
                   $"  日期格式标准化: {DateFormatsStandardized}\n" +
                   $"  数字格式标准化: {NumberFormatsStandardized}\n" +
                   $"  URL标准化: {UrlsStandardized}\n" +
                   $"  标点符号标准化: {PunctuationStandardized}\n" +
                   $"  英文大小写标准化: {EnglishCapitalizationStandardized}\n" +
                   $"  标准化完成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        }
    }

    /// <summary>
    /// 自定义模式替换结果类
    /// </summary>
    public class CustomPatternReplaceResult
    {
        /// <summary>
        /// 模式字典
        /// </summary>
        public Dictionary<string, string> Patterns { get; set; }

        /// <summary>
        /// 是否完成替换
        /// </summary>
        public bool IsCompleted { get; set; }

        /// <summary>
        /// 成功替换的模式
        /// </summary>
        public List<string> SuccessfulReplacements { get; set; } = new List<string>();

        /// <summary>
        /// 失败替换的模式
        /// </summary>
        public List<string> FailedReplacements { get; set; } = new List<string>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成替换结果报告
        /// </summary>
        /// <returns>替换结果报告</returns>
        public string GenerateReport()
        {
            if (!IsCompleted)
            {
                return $"自定义模式替换未完成，错误信息: {ErrorMessage}";
            }

            return $"自定义模式替换报告:\n" +
                   $"  总模式数: {Patterns.Count}\n" +
                   $"  成功替换: {SuccessfulReplacements.Count}\n" +
                   $"  失败替换: {FailedReplacements.Count}\n" +
                   $"  替换完成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        }
    }

    /// <summary>
    /// 批量模式替换结果类
    /// </summary>
    public class BatchPatternReplaceResult
    {
        /// <summary>
        /// 替换列表
        /// </summary>
        public List<PatternReplacement> Replacements { get; set; }

        /// <summary>
        /// 是否完成替换
        /// </summary>
        public bool IsCompleted { get; set; }

        /// <summary>
        /// 替换结果列表
        /// </summary>
        public List<PatternReplaceResult> Results { get; set; } = new List<PatternReplaceResult>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 成功替换的数量
        /// </summary>
        public int SuccessfulReplacements => Results.Count(r => r.IsSuccessful);

        /// <summary>
        /// 失败替换的数量
        /// </summary>
        public int FailedReplacements => Results.Count(r => !r.IsSuccessful);

        /// <summary>
        /// 生成批量替换报告
        /// </summary>
        /// <returns>批量替换报告</returns>
        public string GenerateReport()
        {
            if (!IsCompleted)
            {
                return $"批量模式替换未完成，错误信息: {ErrorMessage}";
            }

            return $"批量模式替换报告:\n" +
                   $"  总替换数: {Replacements.Count}\n" +
                   $"  成功替换: {SuccessfulReplacements}\n" +
                   $"  失败替换: {FailedReplacements}\n" +
                   $"  替换完成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        }
    }
}