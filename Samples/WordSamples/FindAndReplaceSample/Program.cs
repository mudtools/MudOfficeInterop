//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace FindAndReplaceSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 查找和替换示例");

            // 示例2: 替换操作
            Console.WriteLine("\n=== 示例2: 替换操作 ===");
            ReplaceOperationDemo();

            // 示例1: 查找功能详解
            Console.WriteLine("\n=== 示例1: 查找功能详解 ===");
            FindFunctionDemo();

            // 示例3: 格式查找和替换
            Console.WriteLine("\n=== 示例3: 格式查找和替换 ===");
            FormatFindReplaceDemo();

            // 示例4: 正则表达式支持
            Console.WriteLine("\n=== 示例4: 正则表达式支持 ===");
            RegexSupportDemo();

            // 示例5: 高级查找选项
            Console.WriteLine("\n=== 示例5: 高级查找选项 ===");
            AdvancedFindOptionsDemo();

            // 示例6: 批量文本处理
            Console.WriteLine("\n=== 示例6: 批量文本处理 ===");
            BatchTextProcessingDemo();

            // 示例7: 实际应用示例
            Console.WriteLine("\n=== 示例7: 实际应用示例 ===");
            RealWorldApplicationDemo();

            // 示例8: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例8: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 查找功能详解示例
        /// </summary>
        static void FindFunctionDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 添加示例内容
                document.Range().Text = "这是示例文本。\n查找和替换功能演示。\n示例文本包含多个实例。";

                // 获取查找对象
                using var find = document.Range().Find;

                // 基本文本查找
                find.ClearFormatting();
                find.Text = "示例";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindContinue;

                // 执行查找
                bool found = find.Execute();

                if (found)
                {
                    Console.WriteLine("找到了文本 '示例'");
                }
                else
                {
                    Console.WriteLine("未找到文本 '示例'");
                }

                // 查找下一个匹配项
                find.Execute();
                Console.WriteLine("查找下一个匹配项");

                Console.WriteLine("查找功能演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找功能演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 替换操作示例
        /// </summary>
        static void ReplaceOperationDemo()
        {
            try
            {
                string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.dotx");
                string file2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.docx");

                using var app = WordFactory.Open(file);
                using var document = app.ActiveDocument;

                foreach (var section in document.Sections)
                {
                    using var firstPageHeader = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    if (firstPageHeader != null && firstPageHeader.Exists)
                    {
                        ReplaceInRange(firstPageHeader.Range, "$$head$$", "123456");
                    }
                }

                //// 添加示例内容
                //document.Range().Text = "原文本1\n原文本2\n原文本3\n";

                // 获取查找和替换对象
                using var find = document.Range().Find;
                var replace = find; // 替换对象与查找对象是同一个

                // 设置查找和替换参数
                find.ClearFormatting();
                replace.ClearFormatting();

                // 执行替换（只替换第一个匹配项）
                find.Execute(
                    findText: "$$ANFORDNR$$",
                    replaceWith: "123456",
                    replace: WdReplace.wdReplaceOne
                );

                Console.WriteLine("执行单次替换");

                // 执行全部替换
                find.Execute(
                    findText: "$$pubdate$$",
                    replaceWith: DateTime.Now.ToString(),
                    replace: WdReplace.wdReplaceAll
                );

                Console.WriteLine("执行全部替换");

                Console.WriteLine("替换操作演示完成");

                document.SaveAs(file2);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"替换操作演示出错: {ex.Message}");
            }
        }

        private static void ReplaceInRange(IWordRange range, string findText, string replaceText)
        {
            // 使用Find对象进行查找替换
            IWordFind find = range.Find;

            find.ClearFormatting();
            find.Replacement.ClearFormatting();

            find.Text = findText;
            find.Replacement.Text = replaceText;

            // 设置查找选项
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindContinue;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;

            // 执行替换
            var b = find.Execute(replace: WdReplace.wdReplaceAll);
            Console.WriteLine($"替换结果：{b}");
        }

        /// <summary>
        /// 格式查找和替换示例
        /// </summary>
        static void FormatFindReplaceDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;


                // 添加示例内容
                using var range = document.Range();
                range.Text = "普通文本\n粗体文本\n斜体文本\n";

                // 设置粗体文本
                using var boldRange = document.Range(6, 10); // "粗体文本"
                boldRange.Font.Bold = true;

                // 设置斜体文本
                using var italicRange = document.Range(11, 15); // "斜体文本"
                italicRange.Font.Italic = true;

                // 查找粗体文本
                using var find = document.Range().Find;
                find.ClearFormatting();
                find.Font.Bold = true; // 查找粗体文本
                find.Text = ""; // 文本可以为空，只基于格式查找

                // 执行查找
                bool found = find.Execute();
                if (found)
                {
                    Console.WriteLine("找到了粗体文本");
                }

                // 替换粗体格式为下划线格式
                find.ClearFormatting();
                find.Font.Bold = true;
                find.Replacement.ClearFormatting();
                find.Replacement.Font.Underline = WdUnderline.wdUnderlineSingle;

                find.Execute(
                    findText: "",
                    replaceWith: "",
                    replace: WdReplace.wdReplaceAll
                );

                Console.WriteLine("将粗体格式替换为下划线格式");

                Console.WriteLine("格式查找和替换演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式查找和替换演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 正则表达式支持示例
        /// </summary>
        static void RegexSupportDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 添加示例内容
                document.Range().Text = "电话: 138-1234-5678\n邮箱: example@test.com\n日期: 2025-10-06\n";

                // 使用通配符查找电话号码
                using var find = document.Range().Find;
                find.ClearFormatting();
                find.Text = "[0-9]{3}-[0-9]{4}-[0-9]{4}"; // 电话号码模式
                find.MatchWildcards = true;

                bool found = find.Execute();
                if (found)
                {
                    Console.WriteLine("找到了电话号码");
                }

                // 使用通配符查找邮箱
                find.Text = "[a-zA-Z0-9]*@[a-zA-Z0-9]*\\.[a-zA-Z]*";
                find.MatchWildcards = true;

                found = find.Execute();
                if (found)
                {
                    Console.WriteLine("找到了邮箱地址");
                }

                // 使用通配符查找日期
                find.Text = "[0-9]{4}-[0-9]{2}-[0-9]{2}";
                find.MatchWildcards = true;

                found = find.Execute();
                if (found)
                {
                    Console.WriteLine("找到了日期");
                }

                Console.WriteLine("正则表达式支持演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"正则表达式支持演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 高级查找选项示例
        /// </summary>
        static void AdvancedFindOptionsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 添加示例内容
                document.Range().Text = "Word word WORD\nText text TEXT\n";

                using var find = document.Range().Find;
                find.ClearFormatting();

                // 大小写敏感查找
                find.Text = "Word";
                find.MatchCase = true;
                bool found1 = find.Execute();
                Console.WriteLine($"大小写敏感查找: {found1}");

                // 全字匹配查找
                find.Text = "word";
                find.MatchCase = false;
                find.MatchWholeWord = true;
                bool found2 = find.Execute();
                Console.WriteLine($"全字匹配查找: {found2}");

                // 使用同义词库查找
                find.Text = "car";
                find.MatchFuzzy = true;
                bool found3 = find.Execute();
                Console.WriteLine($"同义词查找: {found3}");

                // 向前查找
                find.Text = "word";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                bool found4 = find.Execute();
                Console.WriteLine($"向前查找: {found4}");

                Console.WriteLine("高级查找选项演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"高级查找选项演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 批量文本处理示例
        /// </summary>
        static void BatchTextProcessingDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                using var document = app.ActiveDocument;

                // 添加示例内容
                document.Range().Text = "Mr. Zhang\nMrs. Li\nDr. Wang\nMr. Liu\nMs. Chen\n";

                // 批量替换称谓
                using var find = document.Range().Find;

                // 替换 "Mr." 为 "先生"
                find.Execute(
                    findText: "Mr.",
                    replaceWith: "先生",
                    replace: WdReplace.wdReplaceAll
                );

                // 替换 "Mrs." 为 "夫人"
                find.Execute(
                    findText: "Mrs.",
                    replaceWith: "夫人",
                    replace: WdReplace.wdReplaceAll
                );

                // 替换 "Dr." 为 "博士"
                find.Execute(
                    findText: "Dr.",
                    replaceWith: "博士",
                    replace: WdReplace.wdReplaceAll
                );

                // 替换 "Ms." 为 "女士"
                find.Execute(
                    findText: "Ms.",
                    replaceWith: "女士",
                    replace: WdReplace.wdReplaceAll
                );

                Console.WriteLine("称谓替换完成");

                Console.WriteLine("批量文本处理演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量文本处理演示出错: {ex.Message}");
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

                using var document = app.ActiveDocument;

                Console.WriteLine("开始文档清理...");

                // 添加示例内容
                document.Range().Text = "Mr. Zhang\nMrs. Li\nDr. Wang\nMr. Liu\nMs. Chen\n" +
                    "电话: 138-1234-5678\n邮箱: example@test.com\n日期: 2025-10-06\n" +
                    "某某公司\nXYZ集团\nDEF企业\n";

                // 1. 清理多余的空格
                CleanupExtraSpaces(document);

                // 2. 标准化称谓
                StandardizeTitles(document);

                // 3. 清理空白行
                RemoveExtraBlankLines(document);

                // 4. 标准化日期格式
                StandardizeDateFormats(document);

                // 5. 更新文档属性
                UpdateDocumentProperties(document);

                // 保存清理后的文档
                string filePath = Path.Combine(Path.GetTempPath(), "CleanedDocument.docx");
                document.SaveAs(filePath);

                Console.WriteLine($"文档清理完成: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 清理多余空格
        /// </summary>
        /// <param name="document">文档对象</param>
        private static void CleanupExtraSpaces(IWordDocument document)
        {
            using var find = document.Range().Find;

            // 替换多个空格为单个空格
            find.Execute(
                findText: "  ", // 两个空格
                replaceWith: " ", // 一个空格
                replace: WdReplace.wdReplaceAll
            );

            // 清理行首空格
            find.Execute(
                findText: "^p ", // 段落标记后跟空格
                replaceWith: "^p", // 仅段落标记
                replace: WdReplace.wdReplaceAll
            );

            Console.WriteLine("多余空格清理完成");
        }

        /// <summary>
        /// 标准化称谓
        /// </summary>
        /// <param name="document">文档对象</param>
        private static void StandardizeTitles(IWordDocument document)
        {
            using var find = document.Range().Find;

            // 标准化公司名称
            var companyReplacements = new Dictionary<string, string>
            {
                {"某某公司", "ABC有限公司"},
                {"XYZ集团", "XYZ集团股份有限公司"},
                {"DEF企业", "DEF企业发展有限公司"}
            };

            foreach (var pair in companyReplacements)
            {
                find.Execute(
                    findText: pair.Key,
                    replaceWith: pair.Value,
                    replace: WdReplace.wdReplaceAll
                );
            }

            Console.WriteLine("称谓标准化完成");
        }

        /// <summary>
        /// 删除多余空行
        /// </summary>
        /// <param name="document">文档对象</param>
        private static void RemoveExtraBlankLines(IWordDocument document)
        {
            using var find = document.Range().Find;

            // 删除连续的空行（保留一个）
            find.Execute(
                findText: "^p^p^p", // 三个连续段落标记
                replaceWith: "^p^p", // 两个段落标记
                replace: WdReplace.wdReplaceAll
            );

            // 再次执行以处理更多连续空行
            find.Execute(
                findText: "^p^p^p",
                replaceWith: "^p^p",
                replace: WdReplace.wdReplaceAll
            );

            Console.WriteLine("空白行清理完成");
        }

        /// <summary>
        /// 标准化日期格式
        /// </summary>
        /// <param name="document">文档对象</param>
        private static void StandardizeDateFormats(IWordDocument document)
        {
            using var find = document.Range().Find;

            // 使用通配符查找并标准化日期格式
            find.MatchWildcards = true;

            // 查找 YYYY/MM/DD 格式并替换为 YYYY-MM-DD
            find.Execute(
                findText: "([0-9]{4})/([0-9]{2})/([0-9]{2})",
                replaceWith: "\\1-\\2-\\3",
                replace: WdReplace.wdReplaceAll
            );

            // 查找 YYYY.MM.DD 格式并替换为 YYYY-MM-DD
            find.Execute(
                findText: "([0-9]{4})\\.([0-9]{2})\\.([0-9]{2})",
                replaceWith: "\\1-\\2-\\3",
                replace: WdReplace.wdReplaceAll
            );

            find.MatchWildcards = false;
            Console.WriteLine("日期格式标准化完成");
        }

        /// <summary>
        /// 更新文档属性
        /// </summary>
        /// <param name="document">文档对象</param>
        private static void UpdateDocumentProperties(IWordDocument document)
        {
            // 更新文档属性
            document.Title = "清理后的文档";
            document.Author = "文档清理工具";
            document.Subject = "已清理的文档";
            document.Keywords = "清理, 标准化, 自动化";

            Console.WriteLine("文档属性更新完成");
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                app.Visible = false; // 隐藏Word窗口

                using var document = app.ActiveDocument;

                // 添加示例内容
                document.Range().Text = "这是示例文本。\n查找和替换功能演示。\n示例文本包含多个实例。\n" +
                    "Mr. Zhang\nMrs. Li\nDr. Wang\n" +
                    "电话: 138-1234-5678\n邮箱: example@test.com\n日期: 2025-10-06\n";

                // 创建查找和替换助手
                var findReplaceHelper = new FindAndReplaceHelper(document);

                // 执行基本查找
                bool found = findReplaceHelper.FindText("示例");
                Console.WriteLine($"查找'示例': {found}");

                // 查找所有匹配项
                var positions = findReplaceHelper.FindAllText("示例");
                Console.WriteLine($"找到'示例' {positions.Count} 次");

                // 执行替换
                int replaced = findReplaceHelper.ReplaceAll("示例", "替换后的文本");
                Console.WriteLine($"替换'示例'为'替换后的文本': {replaced} 次");

                // 创建文档清理助手
                var cleanupHelper = new DocumentCleanupHelper(document);

                // 执行完整清理
                var cleanupReport = cleanupHelper.PerformFullCleanup();
                Console.WriteLine(cleanupReport.GenerateSummary());

                // 创建文本搜索管理器
                var searchManager = new TextSearchManager(document);

                // 查找所有邮箱地址
                var emails = searchManager.FindEmailAddresses();
                Console.WriteLine($"找到 {emails.Count} 个邮箱地址");

                // 获取文档统计信息
                var stats = searchManager.GetDocumentStatistics();
                Console.WriteLine(stats.GenerateReport());

                // 创建模式替换助手
                var patternHelper = new PatternReplaceHelper(document);

                // 执行所有模式标准化
                var standardizationReport = patternHelper.PerformAllStandardizations();
                Console.WriteLine(standardizationReport.GenerateReport());

                // 保存文档
                string filePath = Path.Combine(Path.GetTempPath(), "CompleteExampleWithHelpers.docx");
                document.SaveAs(filePath);

                Console.WriteLine($"使用辅助类创建的完整示例文档已保存: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例出错: {ex.Message}");
            }
        }
    }
}