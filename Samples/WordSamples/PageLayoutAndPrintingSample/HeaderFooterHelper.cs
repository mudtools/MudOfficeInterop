//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace PageLayoutAndPrintingSample
{
    /// <summary>
    /// 页眉页脚助手类
    /// </summary>
    public class HeaderFooterHelper
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public HeaderFooterHelper(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 设置首页不同的页眉页脚
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="firstPageHeaderText">首页页眉文本</param>
        /// <param name="firstPageFooterText">首页页脚文本</param>
        /// <param name="otherPagesHeaderText">其他页面页眉文本</param>
        /// <param name="otherPagesFooterText">其他页面页脚文本</param>
        public void SetDifferentFirstPageHeaderFooter(
            int sectionIndex,
            string firstPageHeaderText,
            string firstPageFooterText,
            string otherPagesHeaderText,
            string otherPagesFooterText)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];

                    // 启用首页不同的页眉页脚
                    section.PageSetup.DifferentFirstPageHeaderFooter = 1;

                    // 设置首页页眉
                    if (!string.IsNullOrEmpty(firstPageHeaderText))
                    {
                        using var firstHeaderRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                        firstHeaderRange.Text = firstPageHeaderText;
                    }

                    // 设置首页页脚
                    if (!string.IsNullOrEmpty(firstPageFooterText))
                    {
                        using var firstFooterRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                        firstFooterRange.Text = firstPageFooterText;
                    }

                    // 设置其他页面页眉
                    if (!string.IsNullOrEmpty(otherPagesHeaderText))
                    {
                        using var otherHeaderRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        otherHeaderRange.Text = otherPagesHeaderText;
                    }

                    // 设置其他页面页脚
                    if (!string.IsNullOrEmpty(otherPagesFooterText))
                    {
                        using var otherFooterRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        otherFooterRange.Text = otherPagesFooterText;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置首页不同页眉页脚时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置奇偶页不同的页眉页脚
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="oddHeaderText">奇数页页眉文本</param>
        /// <param name="oddFooterText">奇数页页脚文本</param>
        /// <param name="evenHeaderText">偶数页页眉文本</param>
        /// <param name="evenFooterText">偶数页页脚文本</param>
        public void SetOddAndEvenPagesHeaderFooter(
            int sectionIndex,
            string oddHeaderText,
            string oddFooterText,
            string evenHeaderText,
            string evenFooterText)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];

                    // 启用奇偶页不同的页眉页脚
                    section.PageSetup.OddAndEvenPagesHeaderFooter = 1;

                    // 设置奇数页页眉
                    if (!string.IsNullOrEmpty(oddHeaderText))
                    {
                        using var oddHeaderRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        oddHeaderRange.Text = oddHeaderText;
                    }

                    // 设置奇数页页脚
                    if (!string.IsNullOrEmpty(oddFooterText))
                    {
                        using var oddFooterRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        oddFooterRange.Text = oddFooterText;
                    }

                    // 设置偶数页页眉
                    if (!string.IsNullOrEmpty(evenHeaderText))
                    {
                        using var evenHeaderRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;
                        evenHeaderRange.Text = evenHeaderText;
                    }

                    // 设置偶数页页脚
                    if (!string.IsNullOrEmpty(evenFooterText))
                    {
                        using var evenFooterRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;
                        evenFooterRange.Text = evenFooterText;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置奇偶页不同页眉页脚时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在页眉中添加文本
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="headerType">页眉类型</param>
        /// <param name="text">文本内容</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="alignment">对齐方式</param>
        public void AddHeaderText(
            int sectionIndex,
            WdHeaderFooterIndex headerType,
            string text,
            string fontName = "宋体",
            float fontSize = 12,
            WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphLeft)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var headerRange = section.Headers[headerType].Range;

                    headerRange.Text = text;
                    headerRange.Font.Name = fontName;
                    headerRange.Font.Size = fontSize;
                    headerRange.ParagraphFormat.Alignment = alignment;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加页眉文本时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在页脚中添加文本
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="footerType">页脚类型</param>
        /// <param name="text">文本内容</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="alignment">对齐方式</param>
        public void AddFooterText(
            int sectionIndex,
            WdHeaderFooterIndex footerType,
            string text,
            string fontName = "宋体",
            float fontSize = 12,
            WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphLeft)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var footerRange = section.Footers[footerType].Range;

                    footerRange.Text = text;
                    footerRange.Font.Name = fontName;
                    footerRange.Font.Size = fontSize;
                    footerRange.ParagraphFormat.Alignment = alignment;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加页脚文本时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在页脚中插入页码
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="footerType">页脚类型</param>
        /// <param name="prefix">页码前缀</param>
        /// <param name="suffix">页码后缀</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="alignment">对齐方式</param>
        public void InsertPageNumber(
            int sectionIndex,
            WdHeaderFooterIndex footerType,
            string prefix = "第 ",
            string suffix = " 页",
            string fontName = "宋体",
            float fontSize = 12,
            WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphCenter)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var footerRange = section.Footers[footerType].Range;

                    // 清空原有内容
                    footerRange.Text = "";

                    // 添加前缀
                    if (!string.IsNullOrEmpty(prefix))
                    {
                        footerRange.Text = prefix;
                        footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }

                    // 插入页码字段
                    footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
                    footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // 添加后缀
                    if (!string.IsNullOrEmpty(suffix))
                    {
                        footerRange.Text = suffix;
                    }

                    // 设置格式
                    footerRange.Font.Name = fontName;
                    footerRange.Font.Size = fontSize;
                    footerRange.ParagraphFormat.Alignment = alignment;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入页码时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在页脚中插入总页数
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="footerType">页脚类型</param>
        /// <param name="prefix">总页数前缀</param>
        /// <param name="suffix">总页数后缀</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="alignment">对齐方式</param>
        public void InsertTotalPages(
            int sectionIndex,
            WdHeaderFooterIndex footerType,
            string prefix = "共 ",
            string suffix = " 页",
            string fontName = "宋体",
            float fontSize = 12,
            WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphCenter)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var footerRange = section.Footers[footerType].Range;

                    // 移动到末尾
                    footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // 添加前缀
                    if (!string.IsNullOrEmpty(prefix))
                    {
                        footerRange.Text = prefix;
                        footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }

                    // 插入总页数字段
                    using var field = footerRange.Fields.Add(footerRange, WdFieldType.wdFieldNumPages);
                    footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // 添加后缀
                    if (!string.IsNullOrEmpty(suffix))
                    {
                        footerRange.Text = suffix;
                    }

                    // 设置格式
                    footerRange.Font.Name = fontName;
                    footerRange.Font.Size = fontSize;
                    footerRange.ParagraphFormat.Alignment = alignment;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入总页数时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加包含页码和总页数的页脚
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="footerType">页脚类型</param>
        /// <param name="separator">页码和总页数之间的分隔符</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="alignment">对齐方式</param>
        public void AddPageNumberWithTotal(
            int sectionIndex,
            WdHeaderFooterIndex footerType,
            string separator = " / ",
            string fontName = "宋体",
            float fontSize = 12,
            WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphCenter)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var footerRange = section.Footers[footerType].Range;

                    // 清空原有内容
                    footerRange.Text = "";

                    // 插入当前页码
                    using var field = footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
                    footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // 添加分隔符
                    if (!string.IsNullOrEmpty(separator))
                    {
                        footerRange.Text = separator;
                        footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }

                    // 插入总页数
                    footerRange.Fields.Add(footerRange, WdFieldType.wdFieldNumPages);

                    // 设置格式
                    footerRange.Font.Name = fontName;
                    footerRange.Font.Size = fontSize;
                    footerRange.ParagraphFormat.Alignment = alignment;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加页码和总页数时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加日期到页眉或页脚
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="headerFooterType">页眉/页脚类型</param>
        /// <param name="isHeader">true表示页眉，false表示页脚</param>
        /// <param name="dateFormat">日期格式</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="alignment">对齐方式</param>
        public void AddDate(
            int sectionIndex,
            WdHeaderFooterIndex headerFooterType,
            bool isHeader,
            string dateFormat = "yyyy年MM月dd日",
            string fontName = "宋体",
            float fontSize = 12,
            WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphLeft)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var range = isHeader
                        ? section.Headers[headerFooterType].Range
                        : section.Footers[headerFooterType].Range;

                    // 添加当前日期
                    range.Text = DateTime.Now.ToString(dateFormat);

                    // 设置格式
                    range.Font.Name = fontName;
                    range.Font.Size = fontSize;
                    range.ParagraphFormat.Alignment = alignment;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加日期时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 清除指定节的页眉
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="headerType">页眉类型</param>
        public void ClearHeader(int sectionIndex, WdHeaderFooterIndex headerType)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var headerRange = section.Headers[headerType].Range;
                    headerRange.Text = "";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清除页眉时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 清除指定节的页脚
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="footerType">页脚类型</param>
        public void ClearFooter(int sectionIndex, WdHeaderFooterIndex footerType)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var footerRange = section.Footers[footerType].Range;
                    footerRange.Text = "";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清除页脚时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取页眉内容
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="headerType">页眉类型</param>
        /// <returns>页眉文本内容</returns>
        public string GetHeaderText(int sectionIndex, WdHeaderFooterIndex headerType)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var headerRange = section.Headers[headerType].Range;
                    return headerRange.Text;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取页眉内容时出错: {ex.Message}");
            }

            return "";
        }

        /// <summary>
        /// 获取页脚内容
        /// </summary>
        /// <param name="sectionIndex">节索引（从1开始）</param>
        /// <param name="footerType">页脚类型</param>
        /// <returns>页脚文本内容</returns>
        public string GetFooterText(int sectionIndex, WdHeaderFooterIndex footerType)
        {
            try
            {
                if (sectionIndex > 0 && sectionIndex <= _document.Sections.Count)
                {
                    using var section = _document.Sections[sectionIndex];
                    using var footerRange = section.Footers[footerType].Range;
                    return footerRange.Text;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取页脚内容时出错: {ex.Message}");
            }

            return "";
        }
    }
}