using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindAndReplaceSample
{
    /// <summary>
    /// 查找和替换助手类
    /// </summary>
    public class FindAndReplaceHelper
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public FindAndReplaceHelper(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 基本文本查找
        /// </summary>
        /// <param name="text">要查找的文本</param>
        /// <param name="forward">是否向前查找</param>
        /// <param name="wrap">查找换行方式</param>
        /// <returns>是否找到</returns>
        public bool FindText(string text, bool forward = true, WdFindWrap wrap = WdFindWrap.wdFindContinue)
        {
            try
            {
                var find = _document.Range().Find;
                find.ClearFormatting();
                find.Text = text;
                find.Forward = forward;
                find.Wrap = wrap;

                bool found = find.Execute();
                return found;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找文本时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 查找所有匹配项
        /// </summary>
        /// <param name="text">要查找的文本</param>
        /// <returns>匹配项位置列表</returns>
        public List<Tuple<int, int>> FindAllText(string text)
        {
            var positions = new List<Tuple<int, int>>();

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
                    positions.Add(new Tuple<int, int>(range.Start, range.End));
                    // 移动到下一个位置
                    range = _document.Range(range.End, _document.Content.End);
                    find = range.Find;
                    find.Text = text;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找所有文本时出错: {ex.Message}");
            }

            return positions;
        }

        /// <summary>
        /// 替换第一个匹配项
        /// </summary>
        /// <param name="findText">要查找的文本</param>
        /// <param name="replaceWith">替换文本</param>
        /// <returns>是否替换成功</returns>
        public bool ReplaceFirst(string findText, string replaceWith)
        {
            try
            {
                var find = _document.Range().Find;
                find.ClearFormatting();
                find.Text = findText;
                find.Replacement.ClearFormatting();
                find.Replacement.Text = replaceWith;

                bool replaced = find.Execute(
                    FindText: findText,
                    ReplaceWith: replaceWith,
                    Replace: WdReplace.wdReplaceOne
                );

                return replaced;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"替换第一个匹配项时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 替换所有匹配项
        /// </summary>
        /// <param name="findText">要查找的文本</param>
        /// <param name="replaceWith">替换文本</param>
        /// <returns>替换的次数</returns>
        public int ReplaceAll(string findText, string replaceWith)
        {
            int replaceCount = 0;

            try
            {
                var find = _document.Range().Find;
                find.ClearFormatting();
                find.Text = findText;
                find.Replacement.ClearFormatting();
                find.Replacement.Text = replaceWith;

                // 执行全部替换
                find.Execute(
                    FindText: findText,
                    ReplaceWith: replaceWith,
                    Replace: WdReplace.wdReplaceAll
                );

                // 估算替换次数（Word不直接返回替换次数）
                replaceCount = -1; // 表示执行了全部替换但无法获取确切数量
            }
            catch (Exception ex)
            {
                Console.WriteLine($"替换所有匹配项时出错: {ex.Message}");
            }

            return replaceCount;
        }

        /// <summary>
        /// 基于格式查找
        /// </summary>
        /// <param name="bold">是否粗体</param>
        /// <param name="italic">是否斜体</param>
        /// <param name="underline">下划线类型</param>
        /// <returns>是否找到</returns>
        public bool FindByFormat(bool? bold = null, bool? italic = null, WdUnderline? underline = null)
        {
            try
            {
                var find = _document.Range().Find;
                find.ClearFormatting();

                if (bold.HasValue)
                    find.Font.Bold = bold.Value ? 1 : 0;

                if (italic.HasValue)
                    find.Font.Italic = italic.Value ? 1 : 0;

                if (underline.HasValue)
                    find.Font.Underline = underline.Value;

                find.Text = ""; // 只基于格式查找

                bool found = find.Execute();
                return found;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"基于格式查找时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 替换格式
        /// </summary>
        /// <param name="oldBold">原粗体状态</param>
        /// <param name="newBold">新粗体状态</param>
        /// <param name="oldItalic">原斜体状态</param>
        /// <param name="newItalic">新斜体状态</param>
        /// <returns>是否执行成功</returns>
        public bool ReplaceFormat(bool? oldBold, bool? newBold, bool? oldItalic = null, bool? newItalic = null)
        {
            try
            {
                var find = _document.Range().Find;
                find.ClearFormatting();

                // 设置查找格式
                if (oldBold.HasValue)
                    find.Font.Bold = oldBold.Value ? 1 : 0;

                if (oldItalic.HasValue)
                    find.Font.Italic = oldItalic.Value ? 1 : 0;

                // 设置替换格式
                find.Replacement.ClearFormatting();

                if (newBold.HasValue)
                    find.Replacement.Font.Bold = newBold.Value ? 1 : 0;

                if (newItalic.HasValue)
                    find.Replacement.Font.Italic = newItalic.Value ? 1 : 0;

                find.Text = "";
                find.Replacement.Text = "";

                // 执行全部替换
                find.Execute(
                    FindText: "",
                    ReplaceWith: "",
                    Replace: WdReplace.wdReplaceAll
                );

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"替换格式时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 使用通配符查找
        /// </summary>
        /// <param name="pattern">通配符模式</param>
        /// <returns>是否找到</returns>
        public bool FindWithWildcards(string pattern)
        {
            try
            {
                var find = _document.Range().Find;
                find.ClearFormatting();
                find.Text = pattern;
                find.MatchWildcards = true;

                bool found = find.Execute();
                return found;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用通配符查找时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 使用通配符替换
        /// </summary>
        /// <param name="pattern">通配符模式</param>
        /// <param name="replaceWith">替换文本</param>
        /// <returns>是否执行成功</returns>
        public bool ReplaceWithWildcards(string pattern, string replaceWith)
        {
            try
            {
                var find = _document.Range().Find;
                find.ClearFormatting();
                find.Text = pattern;
                find.Replacement.ClearFormatting();
                find.Replacement.Text = replaceWith;
                find.MatchWildcards = true;

                // 执行全部替换
                find.Execute(
                    FindText: pattern,
                    ReplaceWith: replaceWith,
                    Replace: WdReplace.wdReplaceAll
                );

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用通配符替换时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 高级查找选项
        /// </summary>
        /// <param name="text">查找文本</param>
        /// <param name="matchCase">是否区分大小写</param>
        /// <param name="matchWholeWord">是否全字匹配</param>
        /// <param name="matchFuzzy">是否模糊匹配</param>
        /// <returns>是否找到</returns>
        public bool AdvancedFind(string text, bool matchCase = false, bool matchWholeWord = false, bool matchFuzzy = false)
        {
            try
            {
                var find = _document.Range().Find;
                find.ClearFormatting();
                find.Text = text;
                find.MatchCase = matchCase;
                find.MatchWholeWord = matchWholeWord;
                find.MatchFuzzy = matchFuzzy;

                bool found = find.Execute();
                return found;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"高级查找时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 清除所有查找格式设置
        /// </summary>
        public void ClearFormatting()
        {
            try
            {
                var find = _document.Range().Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清除格式时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取查找对象以进行更详细的自定义
        /// </summary>
        /// <returns>查找对象</returns>
        public IWordFind GetFindObject()
        {
            try
            {
                return _document.Range().Find;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取查找对象时出错: {ex.Message}");
                return null;
            }
        }
    }
}